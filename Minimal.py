from KuliAnalysis import KuliAnalysis

kuli = KuliAnalysis("13")
kuli.kuli_file_name = r"C:\ECS\KULI_130000\data\CoolingSystems\ExCAR_COM.scs"
ok = kuli.initialize()

print("Unit of MeanEffPressure: %s" % (kuli.get_input_com_unit_by_id("MeanEffPressure"))

#kuli.simulate_operating_point(0)
kuli.clean_up()
del kuli