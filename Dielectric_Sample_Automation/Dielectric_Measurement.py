import win32com.client

num_measurements = 5


win32com.client.gencache.EnsureModule('{A7E69DDA-25C0-4373-ADD4-9400E56ECB3A}', 0, 1, 0)

probe_COM = win32com.client.dynamic.Dispatch('Automation8507x.Automation85070.1')
test_return_val = 0
test = probe_COM.Init()
print(test_return_val)
print("Test: ", test)
# Measurement bounds must be set before calibration
probe_COM.SetMeasurement(200e6, 6000e6, 100, 1)
probe_COM.CalibrateProbe()
input("press enter")
probe_COM.TriggerProbe()
input("press enter")
# print(success)
# input("Press Enter to continue")
#
#
#
# for measurement in range(num_measurements):
#     success = probe_COM.TriggerProbe()
#     freq = 0
#     er = 0
#     ei = 0
#     probe_COM.GetMeasurement(5,freq,er,ei)
#     print(freq)
#     print(er)
#     print(ei)
#     print(success)
