import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { CheckboxList, ICheckboxListOption } from '../../../common/controls/CheckboxList';
import type { IAuditPageProps } from './IAuditPageProps';
import styles from './AuditPage.module.scss';

interface IAuditGridItem {
  key: string;
  scope: string;
  location: string;
  status: string;
  details: string;
}

const auditOptions: ICheckboxListOption[] = [
  { key: 'ExpandGroups', label: 'ExpandGroups' },
  { key: 'IncludeListsWithUniquePermissions', label: 'IncludeListsWithUniquePermissions' },
  { key: 'IncludeHiddenListsWithUniquePermissions', label: 'IncludeHiddenListsWithUniquePermissions' },
  { key: 'IncludeListItemsWithUniquePermissions', label: 'IncludeListItemsWithUniquePermissions' }
];

const auditColumns: IColumn[] = [
  {
    key: 'scope',
    name: 'Scope',
    fieldName: 'scope',
    minWidth: 120,
    maxWidth: 180,
    isResizable: true
  },
  {
    key: 'location',
    name: 'Location',
    fieldName: 'location',
    minWidth: 180,
    maxWidth: 320,
    isResizable: true
  },
  {
    key: 'status',
    name: 'Status',
    fieldName: 'status',
    minWidth: 120,
    maxWidth: 160,
    isResizable: true
  },
  {
    key: 'details',
    name: 'Details',
    fieldName: 'details',
    minWidth: 220,
    isResizable: true
  }
];

export const AuditPage: React.FunctionComponent<IAuditPageProps> = (props) => {
  const [selectedOptionKeys, setSelectedOptionKeys] = React.useState<string[]>(
    auditOptions.map((option) => option.key)
  );
  const [auditItems, setAuditItems] = React.useState<IAuditGridItem[]>([]);
  const isSharePointReady: boolean = !!props.sp;

  const onStartAudit = React.useCallback((): void => {
    setAuditItems([]);
  }, []);

  return (
    <section className={styles.auditPage}>
      <div className={styles.header}>
        <h2 className={styles.title}>Permission Audit</h2>
        <p className={styles.subtitle}>Configure the audit scope and review the results.</p>
      </div>

      <CheckboxList
        label="Audit options"
        options={auditOptions}
        selectedKeys={selectedOptionKeys}
        onSelectionChange={setSelectedOptionKeys}
      />

      <Stack horizontal horizontalAlign="start" className={styles.commandBar}>
        <PrimaryButton text="Start" disabled={!isSharePointReady} onClick={onStartAudit} />
      </Stack>

      <div className={styles.grid}>
        <DetailsList
          items={auditItems}
          columns={auditColumns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          ariaLabelForGrid="Permission audit results"
        />
      </div>
    </section>
  );
};
