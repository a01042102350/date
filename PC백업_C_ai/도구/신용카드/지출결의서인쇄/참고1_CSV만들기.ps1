# 참고1.xlsx(1~11월 전표, 아마란스에서 받은 것) → 자동전표 도구가 쓰는 참고 CSV (%TEMP%\참고1_전표.csv)
# 엑셀 없이 읽음(시트 XML이 150MB라 C#으로 한 줄씩 읽음). 사용: . 참고1_CSV만들기.ps1; ConvertTo-JournalCsv $xlsx $csv
# 2026-11-26: 임시 폴더의 참고1_전표.csv가 지워져 '지출결의서 한번에' 5단계(자동전표)가 멈춰서 만듦
if (-not ('JournalXlsxToCsv' -as [type])) {
  Add-Type -ReferencedAssemblies System.IO.Compression, System.IO.Compression.FileSystem, System.Xml -TypeDefinition @'
using System; using System.IO; using System.IO.Compression; using System.Xml; using System.Text; using System.Collections.Generic;
public static class JournalXlsxToCsv {
  static int ColIdx(string r) { int c = 0; foreach (char ch in r) { if (ch >= 'A' && ch <= 'Z') c = c * 26 + (ch - 'A' + 1); else break; } return c; }
  static string Q(string s) { return "\"" + (s ?? "").Replace("\"", "\"\"") + "\""; }
  public static int Convert(string xlsx, string csv, string[] want) {
    using (var z = ZipFile.OpenRead(xlsx)) {
      var ss = new List<string>();
      var se = z.GetEntry("xl/sharedStrings.xml");
      if (se != null) using (var xr = XmlReader.Create(se.Open())) {
        while (xr.Read()) {
          if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "si") {
            var sb = new StringBuilder();
            using (var sub = xr.ReadSubtree()) { while (sub.Read()) { if (sub.NodeType == XmlNodeType.Element && sub.LocalName == "t") sb.Append(sub.ReadElementContentAsString()); } }
            ss.Add(sb.ToString());
          }
        }
      }
      var sh = z.GetEntry("xl/worksheets/sheet1.xml");
      if (sh == null) throw new Exception("참고1.xlsx에 첫 시트가 없습니다");
      int[] cols = null; int n = 0;
      using (var w = new StreamWriter(csv, false, new UTF8Encoding(true)))
      using (var xr = XmlReader.Create(sh.Open())) {
        while (xr.Read()) {
          if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "row") continue;
          var cells = new Dictionary<int, string>(); string curR = null, curT = null;
          using (var sub = xr.ReadSubtree()) {
            while (sub.Read()) {
              if (sub.NodeType != XmlNodeType.Element) continue;
              if (sub.LocalName == "c") { curR = sub.GetAttribute("r"); curT = sub.GetAttribute("t"); }
              else if (sub.LocalName == "v" || sub.LocalName == "t") {
                string v = sub.ReadElementContentAsString(); int k;
                if (curT == "s" && int.TryParse(v, out k) && k >= 0 && k < ss.Count) v = ss[k];
                if (curR != null) cells[ColIdx(curR)] = v;
              }
            }
          }
          if (cols == null) {
            var map = new Dictionary<string, int>(); foreach (var kv in cells) { var h = (kv.Value ?? "").Trim(); if (!map.ContainsKey(h)) map[h] = kv.Key; }
            cols = new int[want.Length]; var miss = new List<string>();
            for (int i = 0; i < want.Length; i++) { int ci; if (map.TryGetValue(want[i], out ci)) cols[i] = ci; else { cols[i] = -1; miss.Add(want[i]); } }
            if (miss.Count > 0) throw new Exception("참고1.xlsx 머리글에 없는 칸: " + string.Join(", ", miss));
            var hd = new string[want.Length]; for (int i = 0; i < want.Length; i++) hd[i] = Q(want[i]); w.WriteLine(string.Join(",", hd));
            continue;
          }
          var o = new string[want.Length];
          for (int i = 0; i < want.Length; i++) {
            string v; if (!cells.TryGetValue(cols[i], out v)) v = "";
            double d;   // 날짜 칸이 숫자(엑셀 날짜)로 들어 있으면 yyyy-MM-dd로
            if ((want[i] == "작성일" || want[i] == "승인일") && double.TryParse(v, out d) && d > 20000 && d < 80000) v = DateTime.FromOADate(d).ToString("yyyy-MM-dd");
            o[i] = Q(v);
          }
          w.WriteLine(string.Join(",", o)); n++;
        }
      }
      return n;
    }
  }
}
'@
}
function ConvertTo-JournalCsv([string]$Xlsx, [string]$Csv) {
  $want = @('승인일', '승인번호', '작성일', '작성번호', '전표라인번호', '계정과목코드', '계정과목명', '차변금액', '대변금액', '적요', '거래처코드', '거래처명', '사업자번호', '프로젝트코드', '프로젝트명', '전표유형명', '품의내역')
  $tmp = Join-Path $env:TEMP ('참고1_읽기_' + [guid]::NewGuid().ToString('N') + '.xlsx')
  Copy-Item -LiteralPath $Xlsx $tmp   # 엑셀에서 열려 있어도 읽을 수 있게 복사본으로
  try { [JournalXlsxToCsv]::Convert($tmp, $Csv, [string[]]$want) } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}
