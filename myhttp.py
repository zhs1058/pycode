import requests, json

from sm4 import sm4_encode, sm4_decode, sm2_encode

if __name__ == '__main__':

    url = "http://127.0.0.1:8080/login"
    data = {
        "data": "admin",
        "sign": "p@ssw3rd"
    }



    new_data = {}

    for key, value in sorted(data.items(), key=lambda x: x[0]): 
        # print(key,value)
        new_data[key] = value
    # print(new_data)

    replace_str = json.dumps(new_data).replace(" ", "")
    print(replace_str)

    key = "2B7A491D64220B09"
    mData = sm4_encode(key, replace_str)
    sign = sm2_encode(replace_str.encode('utf-8'))

    m_data = {'data': mData, 'sign': sign}
    print(m_data)
    response = requests.post(url, m_data)  # json.dumps()将字典解析为字符串
    res = response.json()
    var = res['data']
    print(sm4_decode(key, var))
    #print(response.json())
