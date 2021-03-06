#!/usr/local/bin/python
# coding = utf-8

import uuid
import redis

def create_code(num, length):
#生成”num“个激活码，每个激活码含有”length“位
    result = []
    while True:
        uuid_id = uuid.uuid1()
        temp = str(uuid_id).replace('-', '')[:length]
        if temp not in result:
            result.append(temp)
        if len(result) == num:
            break
    return result


def save_to_redis(num_list):
    r = redis.Redis(host='localhost', port=6379, db=0)
    for code in num_list:
        r.lpush('code',code)

save_to_redis(create_code(200,20))

