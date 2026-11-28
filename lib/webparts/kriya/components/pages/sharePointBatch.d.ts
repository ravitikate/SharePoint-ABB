export interface SharePointBatchResult {
    success: boolean;
    error?: string;
}
export declare const setBackgroundAutoUpdateIndicator: (pageSelector: string, active: boolean) => void;
export declare const createSharePointItemsInBatches: (siteUrl: string, listName: string, digest: string, payloads: Record<string, unknown>[], batchSize?: number, maxConcurrentBatches?: number) => Promise<SharePointBatchResult[]>;
//# sourceMappingURL=sharePointBatch.d.ts.map