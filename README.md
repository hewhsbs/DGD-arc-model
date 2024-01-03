# DGD-arc-model
The process of generating HIAF waveforms using DGD arc models
PSCAD version：4.6.2( X64 ).
GFortran version：4.6.2.
Window : win10.
Spyder: 3.9.2.
Usage:  
1）Fitting Waveform：Put the proposed DGD arc model and the python program mentioned in this article in the folder under the full English path, adjust the range of three variable parameters, start the program, and finally get the optimal parameter combination.
2)  Adjusting parameters to simulate HIAF waveforms:Please put the PSCAD model in the same folder with the white1.out and Guasspink.out files, and then adjust the three variable parameters to generate HIAF waveforms.
Explanation:
1）The DGDM. pscx is the main arc model module built in PSCAD.
2）The white1. out and Guasspink.out are additional noise modules.
3）The DGD1.py is the main program, which includes the dynamic parameter optimization process associated with the PSCAD automation library.
