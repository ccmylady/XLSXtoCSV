import datetime

character_notin_gbk = {'1¼': '1 1/4', '1½': '1 1/2', '1¾': '1 3/4',
                       '¼': '1/4', '½': '1/2', '¾': '3/4',
                       '»': '>>', '«': '<<'}
originstring='CP 06-N½27-N1½ S055-F open'
print(originstring)
for key, value in character_notin_gbk.items():
    originstring=originstring.replace(key,value)

print(originstring)


# 年-月-日 时：分：秒
now_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
print(type(now_time))
print(now_time)

