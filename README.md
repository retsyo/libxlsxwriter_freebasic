# what

This is a [libxlsxwriter, A C library for creating Excel XLSX files](https://github.com/jmcnamara/libxlsxwriter) binding for [FreeBasic](https://www.freebasic.net/). Most of the work is done by [fbfrog](https://github.com/dkl/fbfrog). So that is the reason I call the file as **auto_**xlsxwriter.bi



# how to use

First of all, you should compile libxlsxwriter.dll by yourselves. I don't use Linux/MaxOSX, so please add several lines for them and do test by yourselves if you are in Linux/MaxOSX.

There are many examples, please read them. Most of the examples are translated from the original C examples coming with official [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) by John McNamara, jmcnamara@cpan.org. All of them can be compiled and run without problem, but with only one exception `defined_name_no.bas` which uses complex  C macro so that I don't know how to translate it into [FreeBasic](https://www.freebasic.net/).

But there is still `mini.bas` and `test.bas` which are written by me without using `auto_xlsxwriter.bi`.



# how to do static link

I don't know



# how to split `auto_xlsxwriter.bi` into files

I hope so too for maintenance reason. But I don't know how to do that with [fbfrog](https://github.com/dkl/fbfrog).



# how to write Excel xls file with free library?

I don't know



# how to read Excel xls/xlsx file with free library?

I don't know



# is there a C/FreeBasic library can match [pandas](https://pandas.pydata.org) ?

No. I think there will not be one forever.



# license

FreeBSD license, because [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)









