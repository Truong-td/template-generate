postman request POST 'localhost:8080/api/templates/generate' \
--header 'Content-Type: application/json' \
--body '{
"textData": "{\"application\":{\"name\":\"Digi\"},\"user\":null}",
"tableData": "{\"students\":[{\"name\":\"A\",\"age\":\"0\"},{\"name\":\"B\",\"age\":\"1\"}],\"subjects\":[{\"name\":\"subject A\",\"credit\":\"00\",\"score\":10},{\"name\":\"subject B\",\"credit\":null,\"score\":9}]}",
"flexDataList": [
{
"image": "https://cdn2.fptshop.com.vn/unsafe/Uploads/images/tin-tuc/179088/Originals/meme-noi-tieng-nhat-12.jpg",
"text": "abcdefgh",
"table": [
["A", "B"],
[1, 2]
]
}
]
}'