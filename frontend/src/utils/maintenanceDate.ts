/** Converte dd/mm/aaaa para yyyy-mm-dd (input date). */
export function displayToInputValue(display?: string | null): string {
  const raw = String(display || '').trim()
  if (!raw) {
    return ''
  }

  const brMatch = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
  if (brMatch) {
    return `${brMatch[3]}-${brMatch[2]}-${brMatch[1]}`
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw
  }

  return ''
}

/** Converte yyyy-mm-dd para dd/mm/aaaa (exibição e API). */
export function inputValueToDisplay(input?: string): string {
  const raw = String(input || '').trim()
  if (!raw) {
    return ''
  }

  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (isoMatch) {
    return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}`
  }

  return raw
}
