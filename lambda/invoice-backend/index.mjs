// Lambda invoice-backend（HTTP API invoice-lambda-bridge の裏側）
// 画面（Amplify）からの Gemini 呼び出しを中継する。API キーは Secrets Manager（環境変数 SECRET_ID）にあり、
// このファイルには秘密値を書かないこと（リポは PUBLIC）。
//
// 環境変数
//   SECRET_ID            Secrets Manager のシークレット名（GEMINI_API_KEY などを JSON で保持）
//   TEMPLATE_BUCKET      Excel テンプレートの S3 バケット
//   TEMPLATE_KEY         Excel テンプレートの S3 キー
//   ALLOWED_ORIGINS      任意。呼び出しを許す画面の Origin（カンマ区切り）。既定は本番の Amplify ドメイン
//   ALLOWED_SOURCE_IPS   任意。呼び出しを許す送信元 IP（カンマ区切り）。既定は社内3拠点
//   ALLOWED_MODELS       任意。使ってよい Gemini モデル（カンマ区切り）。既定は gemini-3.1-flash-lite
//   MAX_GEMINI_BODY_BYTES 任意。/api/gemini/generate の本文上限（既定 6000000）
import { SecretsManagerClient, GetSecretValueCommand } from "@aws-sdk/client-secrets-manager";
import { S3Client, GetObjectCommand } from "@aws-sdk/client-s3";

const secretsClient = new SecretsManagerClient({});
const s3Client = new S3Client({});
let cachedSecret = null;

const listEnv = (name, fallback) =>
  String(process.env[name] || fallback)
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);

const ALLOWED_ORIGINS = listEnv("ALLOWED_ORIGINS", "https://master.d10fyz1ex7iu04.amplifyapp.com");
// 社内3拠点。HTTP API では requestContext.http.sourceIp が実際の接続元（X-Forwarded-For は使わない）
const ALLOWED_SOURCE_IPS = listEnv("ALLOWED_SOURCE_IPS", "123.253.152.94,140.82.207.92,202.218.133.135");
const DEFAULT_MODEL = "gemini-3.1-flash-lite";
const ALLOWED_MODELS = listEnv("ALLOWED_MODELS", DEFAULT_MODEL);
const MAX_GEMINI_BODY_BYTES = Number(process.env.MAX_GEMINI_BODY_BYTES || 6000000);
const MAX_LOG_BODY_BYTES = 8192;
const MAX_CONTENT_PARTS = 20; // 文字1 + PDF 画像は最大12ページ
const MAX_OUTPUT_TOKENS = 4096;

export const handler = async (event) => {
  const method = event.requestContext?.http?.method;
  const path = event.rawPath;
  const origin = getHeader(event, "origin");

  if (method === "OPTIONS") {
    return { statusCode: 204, headers: corsHeaders(origin), body: "" };
  }

  // 社内 IP かつ本番画面からの呼び出しだけ受け付ける
  const sourceIp = event.requestContext?.http?.sourceIp || "";
  if (!ALLOWED_SOURCE_IPS.includes("*") && !ALLOWED_SOURCE_IPS.includes(sourceIp)) {
    return json(403, { error: { message: "Forbidden" } }, origin);
  }
  if (!ALLOWED_ORIGINS.includes(origin)) {
    return json(403, { error: { message: "Forbidden" } }, origin);
  }

  if (method === "GET" && path === "/api/health") {
    const secret = await getSecret();
    // GCP のプロジェクト名などの内部情報は返さない
    return json(200, { ok: true, defaultModel: defaultModel(secret) }, origin);
  }

  if (method === "GET" && path === "/api/template/extraction-fields") {
    return getTemplateFromS3(origin);
  }

  if (method === "POST" && path === "/api/gemini/generate") {
    try {
      return await invokeGeminiGenerateContent(event, origin);
    } catch (error) {
      console.error("[gemini.generate.failed]", error);
      return json(500, { error: { message: "Gemini request failed in Lambda" } }, origin);
    }
  }

  if (method === "POST" && path === "/api/log") {
    const body = bodyText(event);
    if (Buffer.byteLength(body, "utf8") > MAX_LOG_BODY_BYTES) {
      return json(413, { error: { message: "Log is too large" } }, origin);
    }
    console.log("[client.log]", body);
    return json(200, { ok: true }, origin);
  }

  return json(404, { error: { message: "Not found" } }, origin);
};

function getHeader(event, name) {
  const headers = event.headers || {};
  for (const [key, value] of Object.entries(headers)) {
    if (key.toLowerCase() === name) return String(value || "");
  }
  return "";
}

function bodyText(event) {
  const body = event.body || "";
  return event.isBase64Encoded ? Buffer.from(body, "base64").toString("utf8") : body;
}

function json(statusCode, body, origin) {
  return {
    statusCode,
    headers: corsHeaders(origin, { "Content-Type": "application/json; charset=utf-8" }),
    body: JSON.stringify(body),
  };
}

function corsHeaders(origin, headers = {}) {
  return {
    ...headers,
    "Access-Control-Allow-Origin": ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0],
    Vary: "Origin",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "content-type",
    "Access-Control-Expose-Headers": "x-template-source,x-template-name",
  };
}

async function getSecret() {
  if (cachedSecret) {
    return cachedSecret;
  }
  const response = await secretsClient.send(new GetSecretValueCommand({ SecretId: process.env.SECRET_ID }));
  cachedSecret = JSON.parse(response.SecretString || "{}");
  return cachedSecret;
}

function defaultModel(secret) {
  const model = normalizeGeminiModelName(secret.GEMINI_DEFAULT_MODEL || DEFAULT_MODEL);
  return ALLOWED_MODELS.includes(model) ? model : ALLOWED_MODELS[0];
}

async function invokeGeminiGenerateContent(event, origin) {
  const raw = bodyText(event);
  if (Buffer.byteLength(raw, "utf8") > MAX_GEMINI_BODY_BYTES) {
    return json(413, { error: { message: "Request is too large" } }, origin);
  }

  let requestJson;
  try {
    requestJson = JSON.parse(raw || "{}");
  } catch {
    return json(400, { error: { message: "Invalid JSON" } }, origin);
  }

  const secret = await getSecret();

  // 画面はクエリ ?model= で送ってくる。許可リスト以外のモデルは使わせない
  const requested = requestJson.model || event.queryStringParameters?.model || "";
  const model = requested ? normalizeGeminiModelName(requested) : defaultModel(secret);
  if (!ALLOWED_MODELS.includes(model)) {
    return json(400, { error: { message: `Model is not allowed: ${model.slice(0, 60)}` } }, origin);
  }

  const contents = requestJson.contents;
  if (!Array.isArray(contents) || contents.length === 0 || contents.length > 4) {
    return json(400, { error: { message: "Invalid contents" } }, origin);
  }
  const partCount = contents.reduce((n, c) => n + (Array.isArray(c?.parts) ? c.parts.length : 0), 0);
  if (partCount === 0 || partCount > MAX_CONTENT_PARTS) {
    return json(400, { error: { message: "Invalid contents" } }, origin);
  }

  const generationConfig =
    requestJson.generationConfig && typeof requestJson.generationConfig === "object"
      ? { ...requestJson.generationConfig }
      : {};
  generationConfig.candidateCount = 1;
  generationConfig.maxOutputTokens = Math.min(Number(generationConfig.maxOutputTokens) || 2400, MAX_OUTPUT_TOKENS);

  const geminiBody = { contents, generationConfig };

  // provider は "generative"（Gemini API キー）のみ使用中
  const apiKey = secret.GEMINI_API_KEY;
  if (!apiKey) {
    return json(500, { error: { message: "Gemini API key is not configured" } }, origin);
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-goog-api-key": apiKey },
    body: JSON.stringify(geminiBody),
  });

  const responseText = await response.text();

  if (!response.ok) {
    const message = extractGeminiErrorMessage(responseText);
    console.error("[gemini.api.error]", response.status, message);
    return json(response.status, { error: { message } }, origin);
  }

  return {
    statusCode: response.status,
    headers: corsHeaders(origin, { "Content-Type": "application/json; charset=utf-8" }),
    body: responseText,
  };
}

function extractGeminiErrorMessage(responseText) {
  if (!responseText) {
    return "Gemini API returned an empty error response.";
  }
  try {
    const payload = JSON.parse(responseText);
    return String(payload.error?.message || "Gemini API error").slice(0, 300);
  } catch {
    return "Gemini API error";
  }
}

function normalizeGeminiModelName(model) {
  return String(model || DEFAULT_MODEL).replace(/^models\//, "").trim();
}

async function getTemplateFromS3(origin) {
  const bucket = process.env.TEMPLATE_BUCKET;
  const key = process.env.TEMPLATE_KEY;

  if (!bucket || !key) {
    return json(500, { error: { message: "Template is not configured" } }, origin);
  }

  const response = await s3Client.send(new GetObjectCommand({ Bucket: bucket, Key: key }));
  const bytes = await response.Body.transformToByteArray();
  const buffer = Buffer.from(bytes);

  return {
    statusCode: 200,
    headers: corsHeaders(origin, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "X-Template-Source": "s3-workbook",
      "X-Template-Name": encodeURIComponent(key.split("/").pop() || key),
    }),
    isBase64Encoded: true,
    body: buffer.toString("base64"),
  };
}
