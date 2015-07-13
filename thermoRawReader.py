
from win32com.client import Dispatch;
from win32com.client import VARIANT as variant;
from pythoncom import *;
import sys;
import matplotlib.pyplot as plt;

fileName = sys.argv[1];

obj = Dispatch('MSFileReader.XRawFile');
obj.open(fileName);

obj.SetCurrentController(0, 1);

numSpec = obj.GetNumSpectra();
print(' ');
print(' ');
print('================');
print('Number of Spectra: ' + str(numSpec));
lowMass = obj.GetLowMass();
highMass = obj.GetHighMass();
print('Mass Range: ' + str(lowMass) + '-' + str(highMass));

pdrt = obj.RTFromScanNum(1);
pdrt1 = obj.RTFromScanNum(numSpec - 1);

# The VARIANT type is very important!
dummyVariant0 = variant(VT_UNKNOWN, 0);
dummyVariant1 = variant(VT_UNKNOWN, []);

dummyVariant2 = variant(VT_EMPTY, []);
dummyVariant3 = variant(VT_UI8, 0);

# Don't ask me why it works. It works...
temp = obj.GetMassListFromRT(pdrt, '', 0, 0, 0, 0, dummyVariant0.value, dummyVariant1, dummyVariant1);
massList = temp[2];
temp = obj.GetChroData(1, 0, 0, '', '', '', 0.0, pdrt, pdrt1, 0, 0, dummyVariant2, dummyVariant2, dummyVariant3.value);
chroList = temp[2];

plt.subplot(211);
plt.plot(massList[0], massList[1]);
plt.xlabel('$Mass (Th)$');
plt.ylabel('$Counts (s^{-1})$');
plt.xlim((lowMass, highMass));
plt.subplot(212);
plt.plot(chroList[0], chroList[1]);
plt.xlabel('$Time (min)$');
plt.ylabel('$Counts (s^{-1})$');
plt.xlim((pdrt, pdrt1));
plt.show();
