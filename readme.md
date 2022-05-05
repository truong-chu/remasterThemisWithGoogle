# Connet Google Drive with Themis

## Cài Đặt

+ PYTHON
  + Cài đặt python 3.7 hoặc cao hơn
  + Chạy cmd để cài phần bổ trợ:
    + ```pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib```
+ Google API
  + Truy cập: https://developers.google.com/sheets/api/quickstart/python
  + Click nút màu xanh ghi "Enable the Google Sheets API"
  + Download `credentials.json` file, copy vào cùng folder với ```codeLoader.py```
    + Có thể sử dụng 2 account khác để download `credentials.json`, lưu ở tên `credentials1.json` và `credentials2.json`. Như vậy sẽ có 300 requests mỗi 100s

## Tùy chỉnh trước khi chạy

+ ```.env.json```
  + ```SHEET_INPUT_ID```: chứa ID của sheet lưu câu trả lời của học sinh
  + ```SHEET_INPUT_NAME```: chứa tên của sheet lưu câu trả lời của học sinh
  + ```SHEET_OUTPUT_ID```: chứa ID của sheet thể hiện bảng điểm
  + ```SHEET_OUTPUT_NAME```: chứa tên của sheet lưu thể hiện bảng điểm
  + ```CONTEST_MODE```: ACM hoặc IOI 
  + ```START_TIME```: Thời gian bắt đầu contest, để tính penalty trong ACM
  + ```WRONG_SUBMISSION_PENALTY```: Thời gian phạt mỗi lần submit sai
+ ```students.json```: Chứ mã bí mật nào ứng với mã học sinh. Mã học sinh không nên viết tiếng việt, hay có dấu cách.

## Chạy

+ ```codeLoader.py```
  + Lệnh: ```python codeLoader.py```
  + Đây là file nhằm download code, lưu tại ```./contestants/```
+ ```scoreUploader.py```
  + Lệnh: ```python scoreUploader.py```
  + Đây là file nhằm đọc Logs từ Themis, update kết quả lên google sheet

## Chú ý

+ ```codeLoader.py``` và ```scoreUploader.py``` là 2 file riêng biệt, ko ảnh hưởng đến nhau. Và phải chạy cả 2 file trong quá trình chấm
+ Nếu 1 trong 2 file ko may crash, thì chỉ cần khởi động lại. Chúng nên tiếp tục hoạt động bình thường. Trong trường hợp không thì có lẽ bạn cần tự debug. Sorry :((
+ Trong ACM mode, 1 bài được tính là AC nếu điểm số >=10. Giả sử có 28 test, có thể đặt điểm 1 test là 0.3575. Đúng 27 test được 9.6525 (Không AC), đúng 28 test được 10.01 (AC)