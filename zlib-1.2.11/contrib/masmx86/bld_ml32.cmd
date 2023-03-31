pushd %2
 %1 /coff /Zi /c /Flmatch686.lst match686.asm
 %1 /coff /Zi /c /Flinffas32.lst inffas32.asm
popd