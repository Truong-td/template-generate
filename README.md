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
---- Test generate template-----

curl --location 'localhost:8080/v1/bpm/internal/generate/template-document' \
--header 'Content-Type: application/json' \
--header 'Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.eyJwcmluY2lwYWwiOiJINHNJQUFBQUFBQUFBSlZSc1U0Q1FSQ2RPeUdTV0NnbW1saG9oWjA1RWxzcVFMUTVEUW5TWUNKWjdoWmMyZHM5ZCtma2FBeWRwWW5HeE1SZjhFK3MvQUEvZ2RyT3VLdmkwUkdubXN5OGVlL056TXNVOGxyQnZsUURUOGVLaVVGZmtZaU9wQnA2bWdhSllqajJBcW1vbDJpcVFvcUVjZTIxVFE3ZjRYeTZzTnlCSWdrQ21RZzhrYUtSeGt6UnNBTnJXYzJYd2RDV05nTFRvUUlaNFhvZXVrd0Y2WEVhK3JCQ0VyeVFScFZSamJEcVg1SnJVazZROFhLTFlzV0hRa3kwTnU1Q2hQV2ZKaWRpVUc2aHRXNzcxcVl3RzF6QkRiaHA3Smd3KysxYXFHZDV2THJrbkFiSXBOQ2x0b2hreVByTWlodit5ZmI5MjkzenBPMENwQXIyRnM5azlhMGFURjdQUDNaY2NIeHdBb1ROT2VzWnJKTEd4azB4WXo1VjFDcS9QelVmSHFlM1owdEcyU0lPLy8rUFV2WDNjdU82akdLaUNNcTVIeG5hVWM3bWhyeTJtSHoyaGJIWFlsSE02WkVpQW1uNEo1RVJtM1Z6U3ZMWnZSRUt6WU51bzlhdEhxY0lnSkNudlM2SnZnRGU2a1hEWndJQUFBPT0iLCJyb2xlcyI6WyJQRF9FQl9BTSJdLCJzdWIiOiJlYl9hbSIsImlzcyI6IkRJR0lMRU5ESU5HIiwiaWF0IjoxNzY3NjA0NzAyLCJleHAiOjE3Njc2MDY1MDJ9.Odmty-AreMaHwFNljQnys7AnTCWRguDSH3IMInWvmZk' \
--data '{
"templateCode": "MB08A_VBPD",
"fileName": "20251231_00016012_MB08A_VBPD",
"textData": {
"toleranceDifference": {
"notHasDifference": true,
"hasDifference": false,
"criteriaDifference": "",
"showProgram": false,
"program": null,
"showProduct": false,
"product": null,
"showUnknown": false,
"differenceUnknown": null,
"differenceLimit": null,
"proposedBasis": []
},
"limitMainTainTime": {
"maintainTime": 12,
"maintainTimeDescription": "kể từ ngày Hợp đồng tín dụng có hiệu lực (ngoại trừ phương thức cấp hạn mức thấu chi hoặc hạn mức cho vay dự phòng có thời hạn hạn mức tối đa 01 năm hoặc các phương thức khác mà pháp luật quy định tối đa 01 năm). Khoản cấp tín dụng này được rà soát ít nhất một năm một lần theo quy định của MSB"
},
"customerInfo": {
"name": "Công ty test luồng soạn thảo",
"customerSegmentationValue": "SME",
"identityNumber": "26122025",
"cif": null,
"businessValue": "Thương mại"
},
"councilMeeting": {
"meetingDate": null,
"meetingTime": null,
"placeCode": null,
"placeText": null,
"meetingEndTime": null
},
"economicSector": {
"sectorLevel5Code": "1150",
"sectorLevel5Value": "Trồng cây thuốc lá, thuốc lào",
"primaryIndustryCode": "NTT5",
"primaryIndustryValue": "FMCG"
},
"initiatorInfo": {
"createdBy": "eb_am",
"createdAt": "2025-12-31 08:49:24"
},
"cashFlowInfo": [
{
"cashFlow": null,
"approvedCashFlow": null,
"isNewCashFlow": false,
"isAdjustCashFlow": false,
"isWithdrawCashFlow": false
}
],
"applicationInfo": {
"segment": "V005",
"solutionPackageValues": "Hạn mức khung",
"approvedAuthorityValue": "Giám đốc quản lý TĐ & PDTD miền",
"submissionPurposeValue": "Cấp mới",
"totalLimit": "8,975,000,000",
"totalLimitText": "Tám tỷ chín trăm bảy mươi lăm triệu đồng",
"riskCategoryValue": "RRTD ≠ 0",
"featureCreditValues": "Hạn mức - Ngắn hạn,Hạn mức - Trung hạn,Hạn mức - Dài hạn",
"templateValue": "Thông báo phê duyệt",
"businessUnitValue": "MSB Hải Phòng",
"areaCode": "NORTH",
"areaValue": "Miền Bắc",
"locationDateText": "Hà Nội, ngày 31 tháng 12 năm 2025",
"branchCode": "020",
"branchName": "CN Hải Phòng"
},
"creditRating": {
"ratingSystem": "MANUAL",
"ratingId": "234343",
"ratingResult": "A",
"policyRank": "B"
}
},
"tableData": {
"creditCouncilsReserveOpinion": [],
"creditCouncils": [],
"program": [],
"products": []
},
"flexData": {
"signDigital": []
}
}'