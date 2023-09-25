
#
# def str_int(src_txt: str = '') -> int | None:
#     if src_txt and isinstance(src_txt, str):
#         try:
#             return int(src_txt)
#         except ValueError:
#             return None
#     return None
#
#
# print(str_int())
# print(str_int('4.4'))


from fastnumbers import fast_real, try_int, try_forceint, try_real

s = "545.2222"
print(fast_real(s))
print(type(fast_real(s)))


s2 = "545.222AA2"
print(try_real(s2))
print(type(try_real(s2)))

print(try_real(s))
print(type(try_real(s)))


print(fast_real("31"))
print(type(fast_real("31")))

print(fast_real("3**1"))

print(try_int('1234'))
print(try_forceint(56.07))
print(try_forceint(56.0))






