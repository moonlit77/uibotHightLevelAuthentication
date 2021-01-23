# -*- coding: utf-8 -*-
import clr
clr.AddReference('LibCommon')
from LibCommon import PyAdapter,ComObject
AdpObj = PyAdapter()
ComObj = ComObject()
RtsCode = 0
field = ''
RtsStr = ''
yxhdbs = ''
ysbty = ''
WiFisz = ''
cpzhmc = ''
SJS = ''

def LogInfo11():
    if AdpObj.GetContinue(11):
        selector = '完成标示选择3'
        AdpObj.LogInfo(selector)

def DataList12():
    if AdpObj.GetContinue(12):
        selector = '{dl}"Action":"Data.List","Title":"数据处理-获取数据清单"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)

def DataSetPara13():
    if AdpObj.GetContinue(13):
        selector = '{dl}"Action":"Data.SetPara","Title":"数据处理-设置参数值","Field":"COLUMN_48"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(13)
        global yxhdbs
        yxhdbs = AdpObj.GetRtsStr()

def LogInfo14():
    if AdpObj.GetContinue(14):
        selector = 'yxhdbs{yxhdbs}'.format(yxhdbs=yxhdbs)
        AdpObj.LogInfo(selector)

def DataSetPara15():
    if AdpObj.GetContinue(15):
        selector = '{dl}"Action":"Data.SetPara","Title":"数据处理-设置参数值","Field":"COLUMN_65"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(15)
        global WiFisz
        WiFisz = AdpObj.GetRtsStr()

def LogInfo16():
    if AdpObj.GetContinue(16):
        selector = 'WiFisz{WiFisz}'.format(WiFisz=WiFisz)
        AdpObj.LogInfo(selector)

def DataSetPara19():
    if AdpObj.GetContinue(19):
        selector = '{dl}"Action":"Data.SetPara","Title":"数据处理-设置参数值","Field":"COLUMN_67"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(19)
        global ysbty
        ysbty = AdpObj.GetRtsStr()

def LogInfo20():
    if AdpObj.GetContinue(20):
        selector = 'ysbty{ysbty}'.format(ysbty=ysbty)
        AdpObj.LogInfo(selector)

def LogInfo22():
    if AdpObj.GetContinue(22):
        selector = '{yxhdbs}-{RtsStr}'.format(yxhdbs=yxhdbs,RtsStr=RtsStr)
        AdpObj.LogInfo(selector)

def LogInfo24():
    if AdpObj.GetContinue(24):
        selector = 'ddd'
        AdpObj.LogInfo(selector)

def MouseAction26():
    if AdpObj.GetContinue(26):
        selector = '{dl}"Action":"Mouse.Action","Title":"鼠标-点击目标","Index":"","App":"","Browser":"ie","Selector":"","Click":"单击","Mouse":"左键","OperType":"后台消息","Timeout":"10000","Delay":"1500"{dr}'.format(dl='{',dr='}')
        AdpObj.RunScript(selector)
        AdpObj.GetRts(26)

def LogInfo28():
    if AdpObj.GetContinue(28):
        selector = 'GZ_RH5GPT299_202001'
        AdpObj.LogInfo(selector)

def DataGetResource29():
    if AdpObj.GetContinue(29):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"YD5G02-013-1-2"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(29)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs30():
    if AdpObj.GetContinue(30):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=500049078]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(30)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource31():
    if AdpObj.GetContinue(31):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"YD5G03-009-1-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(31)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs32():
    if AdpObj.GetContinue(32):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=500054091]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(32)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource34():
    if AdpObj.GetContinue(34):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-731-1-4"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(34)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs35():
    if AdpObj.GetContinue(35):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=500043066]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(35)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource38():
    if AdpObj.GetContinue(38):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"YD4G02-331-2-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(38)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs39():
    if AdpObj.GetContinue(39):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=100096727]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(39)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource42():
    if AdpObj.GetContinue(42):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-690-1-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(42)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs43():
    if AdpObj.GetContinue(43):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=500016050]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(43)
        global bObj
        bObj = AdpObj.GetRtsStr()

def ScriptBlock46():
    if AdpObj.GetContinue(46):
        selector = '{dl}"Action":"Script.Block","Title":"极简标识勾选"{dr}'.format(dl='{',dr='}')
        AdpObj.RunScriptList(selector)

def DataGetResource50():
    if AdpObj.GetContinue(50):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-402-01-13"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(50)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs51():
    if AdpObj.GetContinue(51):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=100018082]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(51)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource53():
    if AdpObj.GetContinue(53):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-660-1-4"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(53)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs54():
    if AdpObj.GetContinue(54):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=100055994]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(54)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource58():
    if AdpObj.GetContinue(58):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-687-1-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(58)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs59():
    if AdpObj.GetContinue(59):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=100096993]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(59)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource62():
    if AdpObj.GetContinue(62):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-690-1-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(62)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs63():
    if AdpObj.GetContinue(63):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=500016050]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(63)
        global bObj
        bObj = AdpObj.GetRtsStr()

def DataGetResource66():
    if AdpObj.GetContinue(66):
        selector = '{dl}"Action":"Data.GetResource","Title":"数据处理-获取资源值","Value":"DM0001-A170-1-1"{dr}'.format(dl='{',dr='}')
        AdpObj.RunDataList(selector)
        AdpObj.GetRts(66)
        global yhbs
        yhbs = AdpObj.GetRtsStr()

def BrowserRunJs67():
    if AdpObj.GetContinue(67):
        selector = '{"Action":"Browser.RunJs","Title":"JS选择标识","Index":"","Browser":"chrome","Jscript":"(function(){$(\'iframe[src*=toBusiPortalMainPage]\').contents().find(\'iframe[src*=toFastSaleMain]\').contents().find(\'iframe[src*=toOfferChoosePage]\').contents().find(\'#resultContainer>div:visible>div:visible>ul>li:visible\').contents().find(\'span[data-offerid=100089616]>input\').click();})()","Timeout":"10000","Delay":"1500"}'
        AdpObj.RunScript(selector)
        AdpObj.GetRts(67)
        global bObj
        bObj = AdpObj.GetRtsStr()

LogInfo11()
DataList12()
DataSetPara13()
LogInfo14()
DataSetPara15()
LogInfo16()
cpzhmc = ComObj.GetValue('Gx_cpzhmc')
SJS = ComObj.GetValue('Gx_SJS')
DataSetPara19()
LogInfo20()
RtsStr = str('GZ_RH5GPT229_202001'.find('RH5GPT299'))
LogInfo22()
if (yxhdbs == '202001'):
    LogInfo24()
MouseAction26()
if (yxhdbs == 'GZ_RH5GPT299_202001'):
    LogInfo28()
    DataGetResource29()
    BrowserRunJs30()
    DataGetResource31()
    BrowserRunJs32()
    if (WiFisz == '299融合套餐专用WiFi6礼包'):
        DataGetResource34()
        BrowserRunJs35()
    if (SJS != '1'):
        DataGetResource38()
        BrowserRunJs39()
    if (cpzhmc == '宽带+IPTV+手机' and ysbty == 'VIP影视包（10元/月）（首年免费、仅限299档次）'):
        DataGetResource42()
        BrowserRunJs43()
        yhbs = 'YD5G03-008-1-2'
        ComObj.SetValue('Gx_ywbs', yhbs)
        ScriptBlock46()
if (yxhdbs == 'GZ_RH4GQG89_202008'):
    DataGetResource50()
    BrowserRunJs51()
    if (cpzhmc != '宽带+IPTV+手机'):
        DataGetResource53()
        BrowserRunJs54()
if (yxhdbs == 'GZ_RH5GCZC129_202001'):
    DataGetResource58()
    BrowserRunJs59()
if (cpzhmc == '宽带+IPTV+手机' and ysbty == 'VIP影视包（10元/月）'):
    DataGetResource62()
    BrowserRunJs63()
elif (cpzhmc == '宽带+IPTV+手机' and ysbty == '90天影视包体验优惠（299档次不可选）'):
    DataGetResource66()
    BrowserRunJs67()

AdpObj.InitAdapter()
