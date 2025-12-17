# OOXML-EXTRACT
Extracts an ooxml document (xlsx, xlsm, vsdx, ...) including the VBA code and prettifies all XML files.
This can then be added to version control, modified, compared, merged, etc.  
It is possible to repackage an OOXML document from it.
However, it is unfortunately not possible to take any modified VBA code into account.

## Usage

### Help
```
ooxml-extract --Help
ooxml-extract extract --Help
ooxml-extract pack --Help
```

### Extract
```
ooxml-extract extract .\Drawing.vsdm -p
```

### Pack
```
ooxml-extract pack .\Drawing .\NewDrawing.vsdm -f
```

### Automerge
```
ooxml-extract automerge .\Original\Stencil.vssm .\Colleague1\Stencil.vssm .\Colleague2\Stencil.vssm .\Merged\Stencil.vssm -f
```