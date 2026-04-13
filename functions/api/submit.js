const TOKEN_REFRESH_BUFFER_MS = 60 * 1000;
const tokenCache = {
  accessToken: null,
  expiresAt: 0,
};

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(),
  });
}

export async function onRequestPost(context) {
  try {
    console.log("START submit");

    const body = await context.request.json();
    validatePayload(body);

    const env = context.env;

    console.log("Getting token...");
    const token = await getAccessToken(env);
    console.log("Got token");

    const company = sanitizeFileName(body.company || "Company");
    const timestamp = formatTimestamp(new Date());
    const newFileName = `${company}_Survey_${timestamp}.xlsx`;

    console.log("Copying template...", newFileName);
    const newItemId = await copyTemplateAndWait({
      token,
      siteId: env.SITE_ID,
      driveId: env.DRIVE_ID,
      templateItemId: env.TEMPLATE_ITEM_ID,
      responsesFolderId: env.RESPONSES_FOLDER_ID,
      newFileName,
    });
    console.log("Template copied. NEW ITEM ID:", newItemId);

    const values = buildInputRangeValues(body);

    console.log("Writing worksheet...", "Inputs_From_User", "E2:G41");
    await updateWorksheetRange({
      token,
      siteId: env.SITE_ID,
      driveId: env.DRIVE_ID,
      itemId: newItemId,
      worksheetName: "Inputs_From_User",
      address: "E2:G41",
      values,
    });
    console.log("Worksheet updated");

    return json({
      ok: true,
      message: "Survey saved successfully",
      fileName: newFileName,
      itemId: newItemId,
    });
  } catch (error) {
    console.error("SUBMIT ERROR:", error);

    return json(
      {
        ok: false,
        error: error instanceof Error ? error.message : "Unknown error",
      },
      500
    );
  }
}

function validatePayload(body) {
  if (!body || typeof body !== "object") {
    throw new Error("Invalid JSON payload");
  }

  if (!body.name || !body.company) {
    throw new Error("Missing required fields: name or company");
  }

  const questionIds = getQuestionIds();
  for (const qid of questionIds) {
    const score = body[`${qid}_score`];
    const evidence = body[`${qid}_evidence`];

    if (score === undefined || evidence === undefined) {
      throw new Error(`Missing score or evidence for ${qid}`);
    }

    if (!Number.isInteger(Number(score)) || Number(score) < 1 || Number(score) > 5) {
      throw new Error(`Invalid score for ${qid}`);
    }

    if (
      !Number.isInteger(Number(evidence)) ||
      Number(evidence) < 0 ||
      Number(evidence) > 2
    ) {
      throw new Error(`Invalid evidence for ${qid}`);
    }
  }
}

async function getAccessToken(env) {
  if (
    tokenCache.accessToken &&
    Date.now() < tokenCache.expiresAt - TOKEN_REFRESH_BUFFER_MS
  ) {
    return tokenCache.accessToken;
  }

  const form = new URLSearchParams();
  form.set("client_id", env.AZURE_CLIENT_ID);
  form.set("client_secret", env.AZURE_CLIENT_SECRET);
  form.set("scope", "https://graph.microsoft.com/.default");
  form.set("grant_type", "client_credentials");

  const res = await fetch(
    `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form.toString(),
    }
  );

  const data = await res.json();

  if (!res.ok || !data.access_token) {
    throw new Error(`Token request failed: ${JSON.stringify(data)}`);
  }

  tokenCache.accessToken = data.access_token;
  tokenCache.expiresAt = Date.now() + (Number(data.expires_in) || 3600) * 1000;

  return data.access_token;
}

async function copyTemplateAndWait({
  token,
  siteId,
  driveId,
  templateItemId,
  responsesFolderId,
  newFileName,
}) {
  const copyUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${templateItemId}/copy`;

  const copyRes = await fetch(copyUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Prefer: "respond-async",
    },
    body: JSON.stringify({
      parentReference: {
        id: responsesFolderId,
      },
      name: newFileName,
    }),
  });

  if (copyRes.status !== 202) {
    const text = await copyRes.text();
    throw new Error(`Template copy failed: ${copyRes.status} ${text}`);
  }

  const monitorUrl =
    copyRes.headers.get("Location") || copyRes.headers.get("operation-location");

  if (!monitorUrl) {
    console.log("No monitor URL returned. Falling back to polling folder.");
    const item = await waitForFileInResponsesFolder({
      token,
      siteId,
      driveId,
      responsesFolderId,
      fileName: newFileName,
    });

    return item.id;
  }

  console.log("Monitor URL received. Polling copy operation...");

  for (let i = 0; i < 12; i++) {
    if (i > 0) {
      await sleep(getPollDelayMs(i));
    }

    const monitorRes = await fetch(monitorUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (monitorRes.status === 202) {
      console.log(`Copy still in progress... attempt ${i + 1}`);
      continue;
    }

    const contentType = monitorRes.headers.get("content-type") || "";

    if (contentType.includes("application/json")) {
      const op = await monitorRes.json();
      console.log("Copy monitor response:", JSON.stringify(op));

      if (op.status === "failed") {
        throw new Error(`Copy operation failed: ${JSON.stringify(op)}`);
      }

      if (op.status === "completed" || op.status === "succeeded") {
        const item = await waitForFileInResponsesFolder({
          token,
          siteId,
          driveId,
          responsesFolderId,
          fileName: newFileName,
        });

        return item.id;
      }
    } else {
      const item = await waitForFileInResponsesFolder({
        token,
        siteId,
        driveId,
        responsesFolderId,
        fileName: newFileName,
      });

      return item.id;
    }
  }

  const item = await waitForFileInResponsesFolder({
    token,
    siteId,
    driveId,
    responsesFolderId,
    fileName: newFileName,
  });

  return item.id;
}

async function waitForFileInResponsesFolder({
  token,
  siteId,
  driveId,
  responsesFolderId,
  fileName,
}) {
  for (let attempt = 0; attempt < 6; attempt++) {
    const item = await findFileInResponsesFolder({
      token,
      siteId,
      driveId,
      responsesFolderId,
      fileName,
    });

    if (item) {
      return item;
    }

    await sleep(getPollDelayMs(attempt));
  }

  throw new Error(`Copied file not found: ${fileName}`);
}

async function findFileInResponsesFolder({
  token,
  siteId,
  driveId,
  responsesFolderId,
  fileName,
}) {
  let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${responsesFolderId}/children`;

  while (url) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    const data = await res.json();

    if (!res.ok) {
      throw new Error(`Failed to list response files: ${JSON.stringify(data)}`);
    }

    const item = (data.value || []).find((x) => x.name === fileName);
    if (item) {
      return item;
    }

    url = data["@odata.nextLink"] || null;
  }

  return null;
}

async function updateWorksheetRange({
  token,
  siteId,
  driveId,
  itemId,
  worksheetName,
  address,
  values,
}) {
  const encodedSheet = encodeURIComponent(worksheetName);

  const url =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}` +
    `/workbook/worksheets/${encodedSheet}/range(address='${address}')`;

  console.log("WORKBOOK URL:", url);

  let lastError;

  for (let attempt = 0; attempt < 3; attempt++) {
    let sessionId = null;

    try {
      sessionId = await createWorkbookSession({ token, siteId, driveId, itemId });

      const res = await fetch(url, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "workbook-session-id": sessionId,
        },
        body: JSON.stringify({ values }),
      });

      const text = await res.text();

      if (res.ok) {
        return;
      }

      lastError = new Error(`Workbook update failed: ${res.status} ${text}`);
      if (!isRetryableStatus(res.status) || attempt === 2) {
        throw lastError;
      }
    } catch (error) {
      lastError = error;
      if (attempt === 2) {
        throw error;
      }
    } finally {
      if (sessionId) {
        await closeWorkbookSession({ token, siteId, driveId, itemId, sessionId });
      }
    }

    await sleep(getPollDelayMs(attempt));
  }

  throw lastError || new Error("Workbook update failed");
}

async function createWorkbookSession({ token, siteId, driveId, itemId }) {
  const url =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}` +
    `/workbook/createSession`;

  console.log("CREATE SESSION URL:", url);

  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ persistChanges: true }),
  });

  const data = await res.json();

  if (!res.ok || !data.id) {
    throw new Error(`Failed to create workbook session: ${JSON.stringify(data)}`);
  }

  return data.id;
}

async function closeWorkbookSession({ token, siteId, driveId, itemId, sessionId }) {
  const url =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}` +
    `/workbook/closeSession`;

  await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "workbook-session-id": sessionId,
    },
  });
}

function buildInputRangeValues(body) {
  return getQuestionIds().map((qid) => [
    body[`${qid}_score`] ?? "",
    body[`${qid}_evidence`] ?? "",
    body[`${qid}_owner`] ?? "",
  ]);
}

function getQuestionIds() {
  return [
    "S1.1", "S1.2", "S1.3", "S1.4", "S1.5",
    "S2.1", "S2.2", "S2.3", "S2.4", "S2.5",
    "S3.1", "S3.2", "S3.3", "S3.4", "S3.5",
    "S4.1", "S4.2", "S4.3", "S4.4", "S4.5",
    "S5.1", "S5.2", "S5.3", "S5.4", "S5.5",
    "S6.1", "S6.2", "S6.3", "S6.4", "S6.5",
    "S7.1", "S7.2", "S7.3", "S7.4", "S7.5",
    "S8.1", "S8.2", "S8.3", "S8.4", "S8.5",
  ];
}

function sanitizeFileName(name) {
  return String(name)
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 80) || "Company";
}

function formatTimestamp(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}_${hh}-${mi}-${ss}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getPollDelayMs(attempt) {
  const delays = [500, 1000, 1500, 2000, 2500, 3000];
  return delays[Math.min(attempt, delays.length - 1)];
}

function isRetryableStatus(status) {
  return [404, 409, 423, 429, 500, 502, 503, 504].includes(status);
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json",
      ...corsHeaders(),
    },
  });
}
