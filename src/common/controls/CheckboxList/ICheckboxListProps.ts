import type { ICheckboxListOption } from './ICheckboxListOption';

export interface ICheckboxListProps {
  options: ICheckboxListOption[];
  selectedKeys: string[];
  label?: string;
  disabled?: boolean;
  onSelectionChange: (selectedKeys: string[]) => void;
}
