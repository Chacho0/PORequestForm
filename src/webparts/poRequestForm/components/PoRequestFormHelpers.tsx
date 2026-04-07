import * as React from 'react';
import styles from './PoRequestForm.module.scss';

export const Section = React.memo((p: { title: string; code?: string; children: React.ReactNode }) => (
  <section className={styles.sectionCard}>
    <div className={styles.sectionHead}>
      <div className={styles.secTitle}>{p.title}</div>
      {p.code ? <div className={styles.secCode}>{p.code}</div> : null}
    </div>
    {p.children}
  </section>
));
Section.displayName = 'Section';

interface StableInputProps {
  id?: string;
  className?: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  disabled?: boolean;
  type?: string;
  readOnly?: boolean;
}

export const StableInput = React.memo<StableInputProps>(({
  id,
  className,
  value,
  onChange,
  placeholder,
  disabled,
  type = 'text',
  readOnly
}) => {
  const handleChange = React.useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      onChange(e.target.value);
    },
    [onChange]
  );

  return (
    <input
      id={id}
      className={className}
      type={type}
      value={value}
      onChange={handleChange}
      placeholder={placeholder}
      disabled={disabled}
      readOnly={readOnly}
    />
  );
});
StableInput.displayName = 'StableInput';

interface StableSelectProps {
  id?: string;
  className?: string;
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
  children: React.ReactNode;
}

export const StableSelect = React.memo<StableSelectProps>(({
  id,
  className,
  value,
  onChange,
  disabled,
  children
}) => {
  const handleChange = React.useCallback(
    (e: React.ChangeEvent<HTMLSelectElement>) => {
      onChange(e.target.value);
    },
    [onChange]
  );

  return (
    <select
      id={id}
      className={className}
      value={value}
      onChange={handleChange}
      disabled={disabled}
    >
      {children}
    </select>
  );
});
StableSelect.displayName = 'StableSelect';

interface StableTextareaProps {
  id?: string;
  className?: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  disabled?: boolean;
  rows?: number;
  readOnly?: boolean;
}

export const StableTextarea = React.memo<StableTextareaProps>(({
  id,
  className,
  value,
  onChange,
  placeholder,
  disabled,
  rows = 3,
  readOnly
}) => {
  const handleChange = React.useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      onChange(e.target.value);
    },
    [onChange]
  );

  return (
    <textarea
      id={id}
      className={className}
      value={value}
      onChange={handleChange}
      placeholder={placeholder}
      disabled={disabled}
      rows={rows}
      readOnly={readOnly}
    />
  );
});
StableTextarea.displayName = 'StableTextarea';
