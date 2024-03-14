import requests, json

from sm4 import sm4_encode, sm4_decode, sm2_encode

if __name__ == '__main__':

    # url = "https://lxapi.lexiangla.com/cgi-bin/token"
    # data = {
    #     "grant_type": "client_credentials",
    #     "app_key": "ee18347e4f2411eaa1b25254002f1020",
    #     "app_secret":"aqmyYu3Q9x6qmDjuehWHsG59r6Sqt8ybk47OB48k"
    # }
    # response = requests.post(url, data)  # json.dumps()将字典解析为字符串
    # res = response.json()
    # print(res)

    unicode_str = '{"data":[{"type":"question","id":"dd5b6b6abe7411ee8f2b367ad2f84d03","attributes":{"title":"\u3010\u79d1\u6280\u5c0f\u8bfe\u5802-\u5bf9\u5916\u7bc7\u30112024.1.29-2024.2.2","summary":"\u4e00\u3001\u4fe1\u606f\u53d1\u5e03\u6a21\u677f \u3010\u5206\u7c7b\u3011\u6807\u9898 \u53d1\u5e03\u5c0f\u7ec4\uff1aXX\u7ec4-\u5e74\u4efd+\u7f16\u53f7\uff08\u56db\u4f4d\uff0c0001\u5f00\u59cb\uff09 \u53d1\u5e03\u4eba\uff1aXX \u53d1\u5e03\u5185\u5bb9\uff1aXX \u4f8b\u5982\uff1a \u3010\u5c0f\u8bfe\u5802\u3011\u7edf\u4e00\u76d1\u63a7\u7cfb\u7edf\u4ecb\u7ecd \u53d1\u5e03\u5c0f\u7ec4\uff1a\u8fd0\u7ef4\u7ec4-20240001 \u53d1\u5e03\u4eba\uff1a\u5f20\u4e09 ...","is_anonymous":0,"read_count":6,"answer_count":0,"concern_count":0,"created_at":"2024-01-29 15:06:03","updated_at":"2024-01-29 15:06:03"},"links":{"platform":"https:\/\/lexiangla.com\/questions\/dd5b6b6abe7411ee8f2b367ad2f84d03?company_from=ee182d6c4f2411eaa5fe5254002f1020"},"relationships":{"owner":{"data":{"type":"staff","id":"20201206260"}},"tags":{"data":[{"type":"tag","id":"bd1043a8926911ee8db4ae3e225648da"}]}}},{"type":"question","id":"b9fbbd46be7411ee975dcacb66bdd3da","attributes":{"title":"\u3010\u79d1\u6280\u5c0f\u8bfe\u5802-\u5bf9\u5185\u7bc7\u30112024.1.29-2024.2.2","summary":"\u4e00\u3001\u4fe1\u606f\u53d1\u5e03\u6a21\u677f \u3010\u5206\u7c7b\u3011\u6807\u9898 \u53d1\u5e03\u5c0f\u7ec4\uff1aXX\u7ec4-\u5e74\u4efd+\u7f16\u53f7\uff08\u56db\u4f4d\uff0c0001\u5f00\u59cb\uff09 \u53d1\u5e03\u4eba\uff1aXX \u53d1\u5e03\u5185\u5bb9\uff1aXX \u4f8b\u5982\uff1a \u3010\u5c0f\u8bfe\u5802\u3011\u7edf\u4e00\u76d1\u63a7\u7cfb\u7edf\u4ecb\u7ecd \u53d1\u5e03\u5c0f\u7ec4\uff1a\u8fd0\u7ef4\u7ec4-20240001 \u53d1\u5e03\u4eba\uff1a\u5f20\u4e09 ...","is_anonymous":0,"read_count":20,"answer_count":3,"concern_count":0,"created_at":"2024-01-29 15:05:04","updated_at":"2024-01-30 09:49:38"},"links":{"platform":"https:\/\/lexiangla.com\/questions\/b9fbbd46be7411ee975dcacb66bdd3da?'  # 例如 "\u4e00\u4e8c\u4e09"

    # 使用unicode_escape解码器对Unicode字符串进行解码
    unicode_str.encode('utf-8').decode('unicode_escape')

    print(unicode_str)


    # new_data = {}
    #
    # for key, value in sorted(data.items(), key=lambda x: x[0]):
    #     # print(key,value)
    #     new_data[key] = value
    # # print(new_data)
    #
    # replace_str = json.dumps(new_data).replace(" ", "")
    # print(replace_str)
    #
    # key = "2B7A491D64220B09"
    # mData = sm4_encode(key, replace_str)
    # sign = sm2_encode(replace_str.encode('utf-8'))
    #
    # m_data = {'data': mData, 'sign': sign}
    # print(m_data)
    # response = requests.post(url, m_data)  # json.dumps()将字典解析为字符串
    # res = response.json()
    # var = res['data']
    # print(sm4_decode(key, var))
    #print(response.json())
