import boto3
from datetime import datetime
from io import BytesIO
import xlsxwriter
import pytz

datetime_strf = datetime.now(pytz.timezone('Asia/Tokyo')).strftime("%Y%m%d%H%M%S")

def write_worksheet(workbook, worksheet_name, headers, data, sheet_title):
# ヘッダーのフォーマット
header_format = workbook.add_format({
"bold": True,
"font_name": "Meiryo UI",
"font_size": 9,
"bg_color":"black",
"font_color":"white",
"align":"center",
"border": 2
})

# セルの基本フォーマット
cell_format = workbook.add_format({"font_name": "Meiryo UI", "font_size":9, "border":1})

# 位置決め
worksheet = workbook.add_worksheet(worksheet_name)

# ヘッダー
for col_num, header in enumerate(headers):
    worksheet.write(2, col_num + 1, header, header_format)    

# データ
for row_num, row_data in enumerate(data, start=3):
    for col_num, value in enumerate(row_data):
        worksheet.write(row_num, col_num + 1, value, cell_format)

# カラム幅の調整
for col_num, header in enumerate(headers):
    max_length = max([len(str(item[col_num])) for item in data] + [len(header)])
    worksheet.set_column(col_num + 1, col_num + 1, max_length + 1)

# フィルター
worksheet.autofilter(2, 1, 2, len(headers))
############################################################

def lambda_handler(event, context):
s3_client = boto3.client("s3")
func_client = boto3.client("lambda")

# Excelファイルに書き込みする処理
function_file_io = BytesIO()
function_workbook = xlsxwriter.Workbook(function_file_io)
############################################################

functions_paginator = func_client.get_paginator("list_functions")

functions = []
for page in functions_paginator.paginate():
    functions.extend(page["Functions"])

# ヘッダーリストの作成
function_headers = ["FunctionName", "Description", "Runtime", "Role", "Name", "Environment", "CostAlloc"]

function_data = []
for function in functions:
    tags = func_client.list_tags(Resource=function["FunctionArn"])
    function_data.append([
        function["FunctionName"],
        function.get("Description", ""),
        function["Runtime"],
        function["Role"],
        tags.get("Tags", {}).get("Name", "None"),
        tags.get("Tags", {}).get("Environment", "None"),
        tags.get("Tags", {}).get("CostAlloc", "None")
    ])

# ワークシートにデータを書き込む
write_worksheet(function_workbook, "Lambda", function_headers, function_data, "AWS Lambda リソース管理表")
############################################################

function_workbook.close()

# S3バケットの指定
s3_bucket = "xxxxxxxxx-my-bucket-name-xxxxxxxxx"

# S3に保存するファイル名の指定&アップロード
function_file_key = f"Resource/Resource_{datetime_strf}/【Lambda】リソース管理表_{datetime_strf}.xlsx"
s3_client.put_object(Bucket=s3_bucket, Key=function_file_key, Body=function_file_io.getvalue())

return {
    "statusCode": 200,
    "body": f"Excelファイルがs3://{s3_bucket}/{function_file_key}に保存されました"
}
