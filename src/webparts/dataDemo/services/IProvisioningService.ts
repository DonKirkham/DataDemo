// ABOUTME: Service contract for provisioning the Conference Events list, content types, and fields.
// ABOUTME: ensureSchema is idempotent; isProvisioned reports whether the target list exists.

export interface IProvisioningSummary {
  created: string[];
  existed: string[];
}

export interface IProvisioningService {
  ensureSchema(): Promise<IProvisioningSummary>;
  isProvisioned(): Promise<boolean>;
}
