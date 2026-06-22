import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import styles from './CheckboxList.module.scss';
import type { ICheckboxListProps } from '.';

export const CheckboxList: React.FunctionComponent<ICheckboxListProps> = (props) => {
  const {
    options,
    selectedKeys,
    label,
    disabled,
    onSelectionChange
  } = props;

  const selectedKeySet: Set<string> = React.useMemo(() => new Set(selectedKeys), [selectedKeys]);

  const onCheckboxChange = React.useCallback((optionKey: string, checked?: boolean): void => {
    const nextSelectedKeys: Set<string> = new Set(selectedKeys);

    if (checked) {
      nextSelectedKeys.add(optionKey);
    } else {
      nextSelectedKeys.delete(optionKey);
    }

    onSelectionChange(Array.from(nextSelectedKeys));
  }, [onSelectionChange, selectedKeys]);

  return (
    <div className={styles.checkboxList}>
      {label && <div className={styles.label}>{label}</div>}
      <div className={styles.options}>
        {options.map((option) => (
          <Checkbox
            key={option.key}
            label={option.label}
            checked={selectedKeySet.has(option.key)}
            disabled={disabled || option.disabled}
            onChange={(_event, checked) => onCheckboxChange(option.key, checked)}
          />
        ))}
      </div>
    </div>
  );
};
