import * as React from 'react';
import styles from './PermissionAudit.module.scss';
import type { IPermissionAuditProps } from './IPermissionAuditProps';
import { AuditPage } from '../AuditPage';

export default class PermissionAudit extends React.Component<IPermissionAuditProps> {
  public render(): React.ReactElement<IPermissionAuditProps> {
    return (
      <section className={`${styles.permissionAudit} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <AuditPage sp={this.props.sp} />
      </section>
    );
  }
}
