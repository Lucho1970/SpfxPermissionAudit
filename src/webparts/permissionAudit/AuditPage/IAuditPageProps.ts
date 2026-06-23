import type { IGraphPermissionAuditService, ISharePointPermissionAuditService } from '../services';

export interface IAuditPageProps {
  graphPermissionAuditService: IGraphPermissionAuditService;
  groupExpansionBatchSize: number;
  groupedViewPreferenceKey: string;
  sharePointPermissionAuditService: ISharePointPermissionAuditService;
}
