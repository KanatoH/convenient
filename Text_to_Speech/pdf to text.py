from spire.pdf import PdfDocument
from spire.pdf import PdfTextExtractOptions
from spire.pdf import PdfTextExtractor

# PdfDocumentクラスのオブジェクトを作成し、PDFファイルをロードします
pdf = PdfDocument()
pdf.LoadFromFile("C:/Users/eeeee/Downloads/convenient/Text_to_Speech/input/CorporateLaw_text3.pdf")

# テキストを保存するための文字列オブジェクトを作成します
extracted_text = ""

# PdfExtractorのオブジェクトを作成します
extract_options = PdfTextExtractOptions()
# シンプルな抽出方法を使用するように設定します
extract_options.IsSimpleExtraction = True

# ドキュメント内のページをループします
for i in range(pdf.Pages.Count):
    # ページを取得します
    page = pdf.Pages.get_Item(i)
    # ページをパラメータとして渡してPdfTextExtractorのオブジェクトを作成します
    text_extractor = PdfTextExtractor(page)
    # ページからテキストを抽出します
    text = text_extractor.ExtractText(extract_options)
    # 抽出されたテキストを文字列オブジェクトに追加します
    extracted_text += text

# 抽出されたテキストをテキストファイルに書き込みます
with open("Text_to_Speech/output/CorporateLaw_text3.txt", "w", encoding="utf-8") as file:
    file.write(extracted_text)
pdf.Close()
