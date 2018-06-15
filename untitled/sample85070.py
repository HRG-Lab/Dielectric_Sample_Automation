import win32com.client

# COM Interface ID's generated using makepy.py
# makepy.py is in C:\ProgramData\Anaconda3\Lib\site-packages\win32com\client
#makepy outputs to C:\Users\ebloom\AppData\Local\Temp\gen_py\3.6
# >>> python makepy.py -i > genfile_test.txt

# Commands from makepy.py Agilent85070 (1.0)
# Output Filename: 7F4CE383-BF50-11D2-9508-9168139D2026x0x1x0.py
# x = win32com.client.gencache.EnsureModule('{7F4CE383-BF50-11D2-9508-9168139D2026}', 0, 1, 0)
# print(x)
# probe_COM = win32com.client.Dispatch('{7F4CE382-BF50-11D2-9508-9168139D2026}')
# print(probe_COM)
# probe_COM.Init()

# Commands from makepy.py Automation8507x 1.0 Type Library (1.0)
# Output Filename: A7E69DDA-25C0-4373-ADD4-9400E56ECB3Ax0x1x0.py
# input('pause')
x = win32com.client.gencache.EnsureModule('{A7E69DDA-25C0-4373-ADD4-9400E56ECB3A}', 0, 1, 0)
print(x)
#input('pause')
# probe_COM = win32com.client.Dispatch('{4CAC772F-2697-4537-B435-1654DF2F8207}')
probe_COM = win32com.client.Dispatch('AUTOMATION8507X.Automation85070')
# print(probe_COM)
probe_COM.Init()

# print(probe_COM)
# print(probe_COM.Init())
