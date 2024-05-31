
all: cstest1.exe cstest2.exe vbtest1.exe vbtest2.exe

cstest1.exe: src\Samples\cs\cstest1.cs src\cs\ComReleaseManager.cs
	csc src\Samples\cs\cstest1.cs src\cs\ComReleaseManager.cs

cstest2.exe: src\Samples\cs\cstest2.cs src\cs\ComReleaseManager.cs
	csc src\Samples\cs\cstest2.cs src\cs\ComReleaseManager.cs

vbtest1.exe: src\Samples\vb\vbtest1.vb src\vb\ComReleaseManager.vb
	vbc src\Samples\vb\vbtest1.vb src\vb\ComReleaseManager.vb

vbtest2.exe: src\Samples\vb\vbtest2.vb src\vb\ComReleaseManager.vb
	vbc src\Samples\vb\vbtest2.vb src\vb\ComReleaseManager.vb

clean:
	del *.bak
	del *.xlsx
	del *.exe
