import requests
import json
class sendRequest():
    def POST(self,url,dict):
        re=requests.post(url,json=dict)
        return re.text



if __name__ == '__main__':

    url="http://221.236.20.217:8093/pcl/services/loanCenter/account/queryPaymentHistory"
    dict={
        "params": {
            "loanNo": "000002017090601542",
            "isPage":1,
            "pageSize":"10",
            "pageNo":"1"
        }
    }
    f=sendRequest().POST(url,dict)

