export interface SharePointBatchResult {
  success: boolean;
  error?: string;
}

export const setBackgroundAutoUpdateIndicator = (pageSelector: string, active: boolean): void => {
  const page = document.querySelector<HTMLElement>(pageSelector);
  if (!page) return;

  let indicator = page.querySelector<HTMLElement>(".kriya-auto-update-indicator");
  if (!active) {
    indicator?.remove();
    return;
  }
  if (indicator) return;

  if (!document.getElementById("kriya-auto-update-indicator-styles")) {
    const styles = document.createElement("style");
    styles.id = "kriya-auto-update-indicator-styles";
    styles.textContent = `
      .kriya-auto-update-indicator {
        position: fixed;
        top: 82px;
        right: 20px;
        z-index: 9998;
        display: inline-flex;
        align-items: center;
        gap: 7px;
        padding: 7px 11px;
        border: 1px solid #c8d8f0;
        border-radius: 16px;
        background: #fff;
        color: #23446f;
        box-shadow: 0 2px 8px rgba(0, 0, 0, .12);
        font: 12px/1.2 Arial, sans-serif;
      }
      .kriya-auto-update-spinner {
        width: 11px;
        height: 11px;
        border: 2px solid #c8d8f0;
        border-top-color: #2563eb;
        border-radius: 50%;
        animation: kriya-auto-update-spin .8s linear infinite;
      }
      @keyframes kriya-auto-update-spin {
        to { transform: rotate(360deg); }
      }
    `;
    document.head.appendChild(styles);
  }

  indicator = document.createElement("div");
  indicator.className = "kriya-auto-update-indicator";
  indicator.setAttribute("role", "status");
  indicator.setAttribute("aria-live", "polite");
  indicator.innerHTML = '<span class="kriya-auto-update-spinner" aria-hidden="true"></span><span>Auto-updating</span>';
  page.appendChild(indicator);
};

const getRetryDelay = (retryAfter: string | null, attempt: number): number => {
  const seconds = retryAfter === null ? NaN : Number(retryAfter);
  const date = retryAfter && Number.isNaN(seconds) ? Date.parse(retryAfter) : NaN;
  if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
  if (Number.isFinite(date)) return Math.max(0, date - Date.now());
  return Math.min(60000, 1000 * (2 ** attempt));
};

export const createSharePointItemsInBatches = async (
  siteUrl: string,
  listName: string,
  digest: string,
  payloads: Record<string, unknown>[],
  batchSize = 20,
  maxConcurrentBatches = 1
): Promise<SharePointBatchResult[]> => {
  const startedAt = Date.now();
  const results: SharePointBatchResult[] = new Array(payloads.length);
  const host = new URL(siteUrl).host;
  const itemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;
  const safeBatchSize = Math.max(1, Math.floor(batchSize));
  const safeBatchConcurrency = Math.max(1, Math.floor(maxConcurrentBatches));

  const batchOffsets = Array.from(
    { length: Math.ceil(payloads.length / safeBatchSize) },
    (_, index) => index * safeBatchSize
  );
  const sendBatch = async (offset: number): Promise<void> => {
    const batchPayloads = payloads.slice(offset, offset + safeBatchSize);
    const batchBoundary = `batch_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const changesetBoundary = `changeset_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const changesetLines: string[] = [];

    batchPayloads.forEach(payload => {
      const body = JSON.stringify(payload);
      changesetLines.push(
        `--${changesetBoundary}`,
        "Content-Type: application/http",
        "Content-Transfer-Encoding: binary",
        "",
        `POST ${itemsUrl} HTTP/1.1`,
        `Host: ${host}`,
        "Accept: application/json;odata=verbose",
        "Content-Type: application/json;odata=verbose",
        `Content-Length: ${new Blob([body]).size}`,
        "",
        body
      );
    });

    changesetLines.push(`--${changesetBoundary}--`);
    const changesetBody = changesetLines.join("\r\n");
    const batchBody = [
      `--${batchBoundary}`,
      `Content-Type: multipart/mixed; boundary="${changesetBoundary}"`,
      `Host: ${host}`,
      `Content-Length: ${new Blob([changesetBody]).size}`,
      "Content-Transfer-Encoding: binary",
      "",
      changesetBody,
      `--${batchBoundary}--`,
      ""
    ].join("\r\n");

    let response: Response | undefined;
    let responseText = "";
    let requestError = "";
    for (let attempt = 0; attempt <= 5; attempt++) {
      try {
        response = await fetch(`${siteUrl}/_api/$batch`, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": `multipart/mixed; boundary="${batchBoundary}"`,
            DataServiceVersion: "3.0",
            MaxDataServiceVersion: "3.0",
            "X-RequestDigest": digest
          },
          body: batchBody
        });
        responseText = await response.text();
      } catch (error) {
        requestError = error instanceof Error ? error.message : String(error);
        break;
      }

      if (response.status !== 429 || attempt === 5) break;
      const waitMs = getRetryDelay(response.headers.get("Retry-After"), attempt);
      console.warn(`SharePoint ${listName} batch throttled; retrying in ${Math.ceil(waitMs / 1000)}s.`);
      await new Promise(resolve => setTimeout(resolve, waitMs));
    }

    const failure = requestError ||
      (response && !response.ok ? `HTTP ${response.status}: ${responseText.slice(0, 250)}` : "");
    if (failure) {
      batchPayloads.forEach((_, index) => {
        results[offset + index] = { success: false, error: failure };
      });
      return;
    }

    const statuses = Array.from(responseText.matchAll(/HTTP\/1\.1\s+(\d{3})[^\r\n]*/gi));
    const batchSucceeded = statuses.length === batchPayloads.length &&
      statuses.every(status => {
        const statusCode = Number(status[1]);
        return statusCode >= 200 && statusCode < 300;
      });

    if (!batchSucceeded) {
      const errorStatuses = statuses
        .map(status => Number(status[1]))
        .filter(status => status < 200 || status >= 300);
      const error = errorStatuses.length
        ? `SharePoint batch item failure (HTTP ${errorStatuses.join(", ")})`
        : "SharePoint batch response did not contain a success status for every item";
      batchPayloads.forEach((_, index) => {
        results[offset + index] = { success: false, error };
      });
      return;
    }

    batchPayloads.forEach((_, index) => {
      results[offset + index] = { success: true };
    });
  };

  for (let index = 0; index < batchOffsets.length; index += safeBatchConcurrency) {
    await Promise.all(
      batchOffsets
        .slice(index, index + safeBatchConcurrency)
        .map(sendBatch)
    );
  }

  console.info(
    `SharePoint ${listName}: ${payloads.length} item(s) in ${batchOffsets.length} batch request(s), ${Date.now() - startedAt}ms.`
  );
  return results;
};
