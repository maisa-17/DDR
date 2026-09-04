param([string]$Path)

try {
  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
  exit 0
}

foreach ($wb in @($excel.Workbooks)) {
  try {
    if ($wb.FullName -ieq $Path) {
      $wb.Close($false)
    }
  } catch {}
}
