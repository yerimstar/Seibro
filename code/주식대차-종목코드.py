import requests
import json
import xmltodict

def Response(code):
    url = "https://seibro.or.kr/websquare/engine/proworks/callServletService.jsp"

    payload = "<reqParam action=\"searchSLBSStockDepthContentList\" task=\"ksd.safe.bip.cmuc.User.process.SearchPTask\"><SECN_TPCD value=\"\"/><INDTP_CLSF_NO value=\"\"/><CUST_SORT_TYPE value=\"\"/><CUST_SORT_TYPE2 value=\"\"/><CALTOT_MART_TPCD value=\""+code+"\"/><FICS_CLSF_NO value=\"\"/></reqParam>\n"
    headers = {
        'Accept': 'application/xml',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,zh;q=0.6,zh-CN;q=0.5,zh-TW;q=0.4,zh-HK;q=0.3',
        'Connection': 'keep-alive',
        'Content-Type': 'application/xml; charset="UTF-8"',
        'Cookie': 'WMONID=sPDJXav32ei; lastAccess=1620358148921; globalDebug=false; JSESSIONID=pctZEm0s5sGIlrg-UwivLZTShXt4iw6NIMaGZaq1MA3vdZKJUk-m!-917870616; SeibroSLBPopup=done; JSESSIONID=bc5Y7wkkH3liLxGsIk2RQdt0MAgFs0Sgtv1iOdhGrGu-uxUHpt2Q!-917870616; WMONID=jzwO4AEb8k2',
        'Host': 'seibro.or.kr',
        'Origin': 'https://seibro.or.kr',
        'Referer': 'https://seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/etc/BIP_CMUC01045P.xml',
        'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="90", "Google Chrome";v="90"',
        'sec-ch-ua-mobile': '?0',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'submissionid': 'P_submission_contentList',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    return response

codeList = ['11','12','13','14','50']
sub_folder = '/Users/yerimstar/PycharmProjects/예탁원/템플릿/'
output_file_name = "주식대차_종목코드.txt"
output_file = open(sub_folder + output_file_name, "w", encoding="utf-8")
output_file.write("{}\t{}\n".format("Name","Code"))
output_file.close()

for code in codeList:
    print(code)
    response = Response(code)
    xmlString = response.text
    jsonDump = json.dumps(xmltodict.parse(xmlString), ensure_ascii=False)
    jsonText = json.loads(jsonDump)
    length = int(jsonText["vector"]["@result"])
    print("length = ",length)
    for i in range(0, length):
        name = str(jsonText["vector"]["data"][i]["result"]["KOR_SECN_NM"]["@value"])
        code = str(jsonText["vector"]["data"][i]["result"]["SHOTN_ISIN"]["@value"])
        output_file = open(sub_folder + output_file_name, "a", encoding="utf-8")
        output_file.write("{}\t{}\n".format(name,code))
        output_file.close()
