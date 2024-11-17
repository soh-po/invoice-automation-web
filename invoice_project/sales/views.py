import os
import uuid
from datetime import datetime
from django.shortcuts import render
from .invoice import create_invoice
from django.conf import settings


def upload_sales_data(request):
    # アップロードファイルを保存する関数
    def handle_uploaded_file(data_path, data):
        with open(data_path, "wb+") as destination:
            for chunk in data.chunks():
                destination.write(chunk)

    def extension_check(files):
        # アップロードファイルの拡張子チェック
        # if type(files) is list: # filesがリストか否か（こちらの書き方はPythonでは非推奨。isinstanceが推奨されている）
        if isinstance(files, list): # filesがリスト型か否か
            for file in files:
                # _, extension = os.path.splitext(file.name) # ファイル名と拡張子に分けて取得
                # result = extension == (".xlsx") # 拡張子が.xlsxかを取得（最後の文字列を比較している）
                result = file.name.endswith(".xlsx") # 拡張子が.xlsxかを取得（最後の文字列を比較している）
                if not result: # Falseであった場合
                    break # ループを終了
        else: # リストでなかった場合(この場合は文字列型)
            # _, extension = os.path.splitext(files.name)
            # result = extension == (".xlsx")
            result = files.name.endswith(".xlsx") # 拡張子が.xlsxかを取得（最後の文字列を比較している）
        return result # TrueかFalseを返す
        
    def upload_file_size_check(files):
        if isinstance(files, list): # filesがリストか
            for file in files:
                result = file.size > 10485760 # ファイルサイズが10MB以上か
                if result: # Trueであった場合
                    break # ループを終了
            # return result
        else: # リストでなかった場合(この場合は文字列型)
            result = files.size > 10485760 # ファイルサイズが10MB以上か
            # return result
        return result # TrueかFalseを返す

    if request.method == "POST":
        now_datetime = datetime.now() # 現在の日時を取得
        #  アップローされたファイルを処理
        sales_datas = request.FILES.getlist("sales_data") # 売上データをリストとして取得
        invoice_template = request.FILES["invoice_template"] # 請求書テンプレートファイルを取得
        # ファイルサイズのバリデーション
        sales_datas_filesize = upload_file_size_check(sales_datas) # 売上データのファイルサイズを確認
        invoice_template_filesize = upload_file_size_check(invoice_template) # 請求書テンプレートのファイルサイズを確認
        if sales_datas_filesize or invoice_template_filesize : # どちらかTrue（ファイルサイズの閾値以上）の場合
            return render(request, "sales/filesize_failed.html", status=400, context={"now_date": now_datetime.strftime("%Y年%m月%d日 %H:%M:%S")}) # 400エラーを返す

        # 拡張子のバリデーション
        sales_data_check = extension_check(sales_datas) # 売上データが.xlsxファイルかどうかを確認
        invoice_template_check = extension_check(invoice_template) # 請求書テンプレートファイルが.xlsxファイルかどうかを確認
        if sales_data_check and invoice_template_check: # 両方ともTrueだった場合
            uuid_dir = str(uuid.uuid4()) # UUIDを生成してディレクトリを分ける
            tmp_dir = os.path.join(settings.MEDIA_ROOT, "tmp", uuid_dir) # アップロードディレクトリのパス
            sales_data_dir = os.path.join(settings.MEDIA_ROOT, "tmp", uuid_dir, "salesbook") # 売上データ　アップロードディレクトリのパス
            create_dates = request.POST.getlist("create_date") # 請求書作成日時をリスト（文字列型）として取得
            os.makedirs(tmp_dir, exist_ok=True) # アップロードディレクトリを作成
            os.makedirs(sales_data_dir, exist_ok=True) # 売上データ　アップロードディレクトリを作成

            # 売上データファイルの保存
            for sales_data in sales_datas: # 売上データをfor文で処理
                sales_data_file = os.path.join(sales_data_dir, sales_data.name) # 保存する売上データファイル名をフルパスで代入
                handle_uploaded_file(sales_data_file, sales_data) # 売上データファイル保存処理
                    
            # 請求書テンプレートファイルの保存
            invoice_template_file = os.path.join(tmp_dir, invoice_template.name) # 保存する売上データファイル名をフルパスで代入
            handle_uploaded_file(invoice_template_file, invoice_template) # 請求書テンプレートファイル保存処理

            # 請求書作成処理
            invoice_url, processing_time = create_invoice(invoice_template_file, tmp_dir, uuid_dir, create_dates)

            # ダウンロード先URLと実行時間、実行日時を返す（実行時間は最終的には表示させない予定）
            return render(request, "sales/invoice_ready.html", {"invoice_url": invoice_url, "processing_time": processing_time, "now_datetime": now_datetime.strftime("%Y年%m月%d日 %H:%M:%S")})
        else:
            return render(request, "sales/ext_failed.html", status=400, context={"now_date": now_datetime.strftime("%Y年%m月%d日 %H:%M:%S")}) # 400エラーを返す
            # return HttpResponse("エクセルファイルではありません。", status=400) # 400エラーを返す

    return render(request, template_name="sales/upload.html")
