import type { IGraphPermissionAuditService, ISharePointPermissionAuditService } from '../services';

export interface IAuditPageProps {
  graphPermissionAuditService: IGraphPermissionAuditService;
  groupedViewPreferenceKey: string;
  sharePointPermissionAuditService: ISharePointPermissionAuditService;
}
