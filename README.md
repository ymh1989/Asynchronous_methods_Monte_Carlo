##Asynchronous methods for Monte Carlo##

This repo contains an implementation of `Microsoft Excel addin` by using `C#`. Especially, I employ the [Excel-DNA](https://exceldna.codeplex.com/) which is an independent project to integrate .NET into Excel. If you would like to complie the code, please follow `Getting Started` in this [link](https://exceldna.codeplex.com/).

###Environment###
- CPU : Intel(R) Core(TM) i5-6400 @ 2.7GHZ 
- RAM : DDR3L 16GB PC3-12800
- Microsoft Visual Studio Community 2013, Microsoft Excel 2013 (64bit)

###Procedure###
1. Run `MC_test_async-AddIn64-packed.xll`.
2. Open `async_test.xlsm` or `parallelfor_test.xlsm`.
3. Click `Calculate`.

** This program is complied for `x64`. If you want to run for `x86`, please complie in `x86` mode. Detailed information for compling with `Excel-DNA` is described in `Init_Korean.pdf`.

###Result###
- Please check out `parallelfor_async_test.gif` in `code` directory.
- By using `parallel for`, the computational cost of Monte Carlo simlation(MCS) for pricing derivatives decreases significantly. The main reason is that `parallel for` uses all of cpu resources with threads. You can confirm this in `parallelfor_test.xlsm`
- `Excel-DNA` provides `ExcelAsyncUtil.Run` which is powerful things to implement asynchronous functions. This does not makes `freezing` of `Excel` window despite of a sort of MCS.

###Future work###
- Real Time data(RTD) by using asyncronous programming