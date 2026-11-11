const PRIVATE_EAN_PREFIX = '29'

function ean13CheckDigit(base12: string) {
  const sum = base12.split('').reduce((total, digit, index) => total + Number(digit) * (index % 2 === 0 ? 1 : 3), 0)
  return String((10 - (sum % 10)) % 10)
}

/** 產生以 29 開頭、含 EAN-13 檢查碼的內部條碼，並避開目前已使用的條碼。 */
export function generateUniqueToolBarcode(existingBarcodes: string[]) {
  const existing = new Set(existingBarcodes.map((value) => value.trim()).filter(Boolean))
  const seed = BigInt(Date.now())

  for (let attempt = 0; attempt < 1000; attempt += 1) {
    const serial = ((seed + BigInt(attempt)) % 10_000_000_000n).toString().padStart(10, '0')
    const base12 = `${PRIVATE_EAN_PREFIX}${serial}`
    const barcode = `${base12}${ean13CheckDigit(base12)}`
    if (!existing.has(barcode)) return barcode
  }

  throw new Error('暫時無法產生未使用的條碼，請稍後再試')
}

/** 依條碼內容產生可顯示、下載或嵌入 Excel 的 QR Code PNG。 */
export async function createQrCodeDataUrl(value: string, width = 280) {
  const QRCode = await import('qrcode')
  return QRCode.toDataURL(value, {
    errorCorrectionLevel: 'M',
    margin: 1,
    width,
    color: { dark: '#000000', light: '#ffffff' },
  })
}
