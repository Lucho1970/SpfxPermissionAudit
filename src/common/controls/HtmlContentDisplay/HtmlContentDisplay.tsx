import * as React from 'react';
import styles from './HtmlContentDisplay.module.scss';
import type { IHtmlContentDisplayProps } from '.';

export const HtmlContentDisplay: React.FunctionComponent<IHtmlContentDisplayProps> = (props) => {
  const {
    html,
    height,
    ariaLabel,
    className
  } = props;

  return (
    <div
      aria-label={ariaLabel}
      className={`${styles.htmlContentDisplay} ${className || ''}`}
      role="region"
      style={{ height }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
};
