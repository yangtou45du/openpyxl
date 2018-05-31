import requests
class sendRequest():
    def POST(self,url,dict):
        re=requests.post(url,data=dict)
        print re.text
if __name__ == '__main__':
    dict={'action':'com.hansy.system.login.action.SystemInitAction','method':'serverLoginLog'}
    url="http://221.236.20.215:802/ams/ContextEngine"
    f=sendRequest().POST(url,dict)