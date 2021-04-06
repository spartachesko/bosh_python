import pymcprotocol

# If you use Q series PLC
pymc3e = pymcprotocol.Type3E()
# if you use L series PLC,
# pymc3e = pymcprotocol.Type3E(plctype="L")
# if you use QnA series PLC,
# pymc3e = pymcprotocol.Type3E(plctype="QnA")
# if you use iQ-L series PLC,
# pymc3e = pymcprotocol.Type3E(plctype="iQ-L")
# if you use iQ-R series PLC,
# pymc3e = pymcprotocol.Type3E(plctype="iQ-R")

# If you use 4E type
# pymc4e = pymcprotocol.Type4E()

# If you use ascii byte communication, (Default is "binary")
pymc3e.setaccessopt(commtype="binary")
pymc3e.connect("192.168.3.39", 1025)

# read from D2045 to D2055
wordunits_values = pymc3e.batchread_wordunits(headdevice="D2045", readsize=10)
print(wordunits_values)
