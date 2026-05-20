interface ReportCheckboxProps {
  checked: boolean
  onChange: (checked: boolean) => void
  title: string
  description: string
}

export function ReportCheckbox({
  checked,
  onChange,
  title,
  description,
}: ReportCheckboxProps) {
  return (
    <label className="checkbox">
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.target.checked)}
      />
      <span className="checkbox-box" aria-hidden="true" />
      <span className="checkbox-content">
        <span className="checkbox-title">{title}</span>
        <span className="checkbox-description">{description}</span>
      </span>
    </label>
  )
}
