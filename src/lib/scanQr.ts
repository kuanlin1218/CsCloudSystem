import { prepareZXingModule, readBarcodes } from 'zxing-wasm/reader'
import readerWasmUrl from 'zxing-wasm/reader/zxing_reader.wasm?url'

// 刀具掃碼視窗｜使用網站建置附帶的 WebAssembly 檔案，適用 GitHub Pages 子路徑。
prepareZXingModule({ overrides: { locateFile: () => readerWasmUrl } })

export async function scanQr(image: ImageData): Promise<string> {
  const [result] = await readBarcodes(image, {
    formats: ['QRCode'],
    tryHarder: true,
    tryDownscale: true,
    tryInvert: true,
    maxNumberOfSymbols: 1,
  })
  return result?.text.trim() || ''
}
