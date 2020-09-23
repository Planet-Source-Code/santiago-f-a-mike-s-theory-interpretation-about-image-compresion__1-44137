Attribute VB_Name = "modMSCFile"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Type to save the Repeated Block position
Public Type tClonePos
    X As Integer
    Y As Integer
End Type

' Type with the information of a Repeated Block position
Public Type tCloneBlock
    Info() As tClonePos
    MasterID As Byte
End Type

' Type with the information of a Master Block
Public Type tMasterBlock
    Info() As Byte
    X As Integer
    Y As Integer
End Type

' Type with a information of a pixel color
Public Type tPixelInfo
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

' Tipo para usar en el encabezado del archivo
' para indicar cosas como tipo de archivo,
' blockHeight, blockWidth, Masters, etc
' This type is used for the header of the compresed
' file. It will indicates the Block Height, Width,
' how many master there are, and more info
Public Type tHeader
    FileType As String * 3
    BlockWidth As Byte
    BlockHeight As Byte
    TotalMasters As Byte
    ImgWidth As Integer
    ImgHeight As Integer
End Type

' Tipo para indicar la cantidad de clones de un master y alguna otra info
' Tyoe to indicate the amount of repeated blocks (clones) of a Master
Public Type tPreClone
    Cant As Integer
End Type

' Indica el tamaño en bytes del encabezado
' Indicate the size of the file header
Private Const HeaderSize = 10
' Indica el porcetage de parecidos que tienen que ser los bloques
' It indicates to the algorithm the amount of "similar" must be a clone block from his master
Dim ErrorPercent As Byte
' Indica si se quiere que se vean los shapes del main para ver como se procesa
' It indicates if you wanna see the progress of the algorithm
Dim ShowBlockToVisualHelp As Boolean
' Sirve para tener un control de todos los pixeles usados
' It is used to know who pixels where used for a Master or Clone block
Dim pUsed() As Boolean
' Guardo todos los bytes del PictureBox
' It Matrix is used for save in memory the pixels information of the image
Dim bPixels() As Byte
' Guardo globalmente los valores del picture box
' I save the Image info in globlas variables
Dim picWidth As Long
Dim picHeight As Long
Dim BlockWidth As Byte
Dim BlockHeight As Byte

' Con este procedimiento se graba en formato MSC
' With this procedure you can save a image of a picture box into the MSC format
' Parameters:
'       Pic: The picture box that contain the image to been save
'       Dest: The path of destination
'       Error: It indicates to the algorithm the amount of "similar" must be a clone block from his master
'       ShowProgres: Indicate if you wanna see the progress in the frmMain (slower)
Public Sub SaveAsMSC(Pic As PictureBox, Dest As String, Error As Byte, ShowProgres As Boolean)
    Dim lTime As Long
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    Dim l As Integer
    
    ' Controlo el tiempo que se tarda
    ' I get the start time to know how long that procedure takes
    lTime = GetTickCount()
    
    ' Guardo la forma de actuar
    ' Save to the globals parameters
    ErrorPercent = Error
    ShowBlockToVisualHelp = ShowProgres
    
    ' Tomo todos los pixels de la imagen
    ' With this function a save all pixel information to the bPixels matrix
    GetDIs Pic, bPixels(), picWidth, picHeight
    
    ' Save the real pictureBox size
    picWidth = picWidth + 1
    picHeight = picHeight + 1

    ' I set the similar Block width and height
    BlockWidth = 8: BlockHeight = 8
    
    Dim xM As Integer, yM As Integer
    
    ' Guardo cada bloque maestro
    ' Here i will save every Master Block Info
    Dim MasterBlocks() As tMasterBlock
    ' Bloque para realizar la comparación
    ' It is used to make a comparation (save a block that perhaps will be a master block)
    Dim ProbableMaster() As Byte
    ' Guardo la pos. de cada clon
    ' Save the all clones position
    Dim CloneBlocks() As tCloneBlock
    
    ' Redimensiono la matriz que va a contener todos los bloques maestros
    ' Redim the masterblocks matrix
    ReDim MasterBlocks(0)          ' --> MasterBlocks (MasterID)
    ' Redimensiono el probableMaster
    ' Redim the ProbableMaster block --> ProbableMaster (BlockX,BlockY,BlockRGB)
    ' BlockRGB --> R=2; G=1; B=0;
    ReDim ProbableMaster(BlockWidth - 1, BlockHeight - 1, 2) As Byte
    ' Redimensiono los bloques clones
    ' Redim the clones block info
    ReDim CloneBlocks(0) As tCloneBlock  ' CloneBlocks(MasterID).Info(NumberOfThisMaster).x/y
    ' Redimensiono la matriz para el control de pixeles
    ' Redim the pUsed matrix used to save who pixels were used for a Master Block or a clone
    ' pUsed(PicX,PicY)
    ReDim pUsed(Pic.ScaleWidth, Pic.ScaleHeight) As Boolean

    ' Voy recorriendo la imagen tomando los bloques maestros
    ' Run trought the image getting block by 8 x 8 pixels to see if it have a clone
    For yM = 0 To picHeight - BlockHeight Step BlockHeight
        For xM = 0 To picWidth - BlockWidth Step BlockWidth
            
            ' Only for guide in frmMain (Visual Help)
            If ShowBlockToVisualHelp = True Then
                frmMain.shpMaster.Move xM, yM
                DoEvents
            End If
            
            Dim ClonesPos() As tClonePos

            ' Si la posición a analizar no fue usada con anterioridad...
            ' If in pos (xM; yM) there is a free pixel... (a free pixel is a pixel that is not used in a Master Block or a Clone Block)
            If NotOverlayed(xM, yM) = True Then
                ' Copio un pedazo de la imagen en una matriz de BlockWidth x BlockHeight para despues comparar con esta
                ' I will copy a piece of image to compare if there is another similar block to determinate if the fisrt block can be a master
                If CopyBlock(bPixels(), xM, yM, BlockWidth, BlockHeight, ProbableMaster()) = True Then
                    ' Veo si el pedazo que copie sirve para un MasterBlock
                    ' Now, i check if there is a similar block in the image like the fisrt a copied before in ProbableMaster matrix
                    If ViewClones(ProbableMaster(), xM, yM, ErrorPercent, ClonesPos()) = True Then
                        ' Si hubo bloques similares...
                        ' If were similars blocks
                        ' Imprimo la cantidad de bloques similares
                        ' Print the amount of clones blocks
                        Debug.Print UBound(ClonesPos())
                        ' Redimensiono en 1 la cantidad de bloques maestros
                        ' Add one more space in the MasterBlocks matrix
                        ReDim Preserve MasterBlocks(UBound(MasterBlocks) + 1) As tMasterBlock
                        ReDim MasterBlocks(UBound(MasterBlocks)).Info(BlockWidth - 1, BlockHeight - 1, 2) As Byte
                        ' Guardo la info del Master
                        ' Save the master block info
                        MasterBlocks(UBound(MasterBlocks)).Info() = ProbableMaster()
                        MasterBlocks(UBound(MasterBlocks)).X = xM
                        MasterBlocks(UBound(MasterBlocks)).Y = yM
                        ' Guardo la info de cada clon de ese master
                        ' Now, save the clones information from this master
                        ReDim Preserve CloneBlocks(UBound(CloneBlocks) + 1) As tCloneBlock
                        ReDim CloneBlocks(UBound(CloneBlocks)).Info(UBound(ClonesPos())) As tClonePos

                        CloneBlocks(UBound(CloneBlocks)).Info() = ClonesPos()

                        ' Marco los bits usados por los clones y el master
                        ' Save who pixels were affected by the news master and clones
                        ' I used FillMemory to decrease the processing time
                        ' But it can be done with a simple pUsed(x,y)=true for every pixel used in a clon and a master block
                        For j = yM To yM + BlockHeight - 1
                            FillMemory ByVal VarPtr(pUsed(xM, j)), BlockWidth * 2, ByVal 255
                        Next j
                        For k = 1 To UBound(CloneBlocks(UBound(CloneBlocks())).Info())
                            For j = CloneBlocks(UBound(CloneBlocks())).Info(k).Y To CloneBlocks(UBound(CloneBlocks())).Info(k).Y + BlockHeight - 1
                                FillMemory ByVal VarPtr(pUsed(CloneBlocks(UBound(CloneBlocks())).Info(k).X, j)), BlockWidth * 2, ByVal 255
                            Next j
                        Next k

                    End If
                End If
            End If
            
        Next xM
    Next yM
    
    ' Informa el resultado
    Dim Clones As Long
    Dim WithCompres As Long
    
    ' I recount the totals clones used to calculate the final file size
    For i = 1 To UBound(CloneBlocks())
        Clones = Clones + UBound(CloneBlocks(i).Info())
    Next i
    
    ' Master Block:    BlockWidth x BlockHeight x 3 bytes    ' RGB Info
    '               +   4 bytes                              ' Posición
    '               +  26 bytes                              ' Agregados por VB cuando se graba un typo del usuario (Added by VB when you save a type, i think)
    '               ------------
    '                 222 bytes
    
    ' Clone Block:     4 bytes                               ' Posición
    '               +  0 byte                                ' Master ID (No lo uso)
    '               ------------
    '                  0 bytes
    
    ' Pixels Libres (FreePixels):  Totales (Totals) - Usados en clones y master (Used in masters and clone)
    ' Bytes per FreePixel: Pixels Libres * 3
    
    ' Calculate the image size with compress
    WithCompres = (UBound(MasterBlocks())) * (BlockWidth * BlockHeight * 3 + 4 + 26) _
                + Clones * 4 _
                + (((Pic.ScaleWidth) * (Pic.ScaleHeight)) - (UBound(MasterBlocks())) * (BlockWidth * BlockHeight) - Clones * (BlockWidth * BlockHeight)) * 3 _
                + HeaderSize _
                + 2 * UBound(MasterBlocks())
    
    ' Info the result
    MsgBox "Total Masters: " & UBound(MasterBlocks()) & vbCrLf & _
           "Total Clones: " & Clones & vbCrLf & _
           "Imagen sin Compresión: " & Round(Pic.ScaleWidth * Pic.ScaleHeight * 3 / 1024, 2) & " KB" & vbCrLf & _
           "Imagen con compresión: " & Round(WithCompres / 1024, 2) & " KB" & vbCrLf & _
           "Porcentage compreso: " & Round(100 - ((WithCompres / 1024) * 100 / (Pic.ScaleWidth * Pic.ScaleHeight * 3 / 1024)), 2) & " %" & vbCrLf & _
           "Tiempo: " & GetTickCount - lTime & " ms", vbInformation
           
    ' Now in english
    MsgBox "Total Masters: " & UBound(MasterBlocks()) & vbCrLf & _
           "Total Clones: " & Clones & vbCrLf & _
           "Image without Compresion: " & Round(Pic.ScaleWidth * Pic.ScaleHeight * 3 / 1024, 2) & " KB" & vbCrLf & _
           "Image with compresion: " & Round(WithCompres / 1024, 2) & " KB" & vbCrLf & _
           "Compressed: " & Round(100 - ((WithCompres / 1024) * 100 / (Pic.ScaleWidth * Pic.ScaleHeight * 3 / 1024)), 2) & " %" & vbCrLf & _
           "Time: " & GetTickCount - lTime & " ms", vbInformation
           
    ' Informo el encabezado
    Dim cHeader As tHeader
    ' I save the data will be saved into the file header
    With cHeader
        .BlockHeight = BlockHeight
        .BlockWidth = BlockWidth
        .FileType = "MSC"
        .ImgHeight = picHeight
        .ImgWidth = picWidth
        .TotalMasters = UBound(MasterBlocks())
    End With
    
    ' Call to the SaveTheFile rutine to save the file
    Call SaveTheFile(Dest, MasterBlocks(), CloneBlocks(), cHeader)
        
    ' Saco de memoria los pixels usados
    ' Set to nothing all varibales to unload of memory
    ReDim bPixels(0, 0, 0)
End Sub

' Copio una matriz a otra
' It function will copy a matrix to another one (i used to copy a piece of image to a 8x8 matrix)
' Parameters:
'       Origen(): A matrix like m(X,Y,2)
'       xOffset and yOffset: indicates where to start the copy
'       Width and Height: Indicates the final dimension of the destination matrix
'       Dest(): A matrix like m(Width,Height,2-->Used for RGB) that will be the destination
Private Function CopyBlock(Origen() As Byte, xOffset As Integer, yOffset As Integer, Width As Byte, Height As Byte, Dest() As Byte) As Boolean
On Error GoTo Solucion
    Dim i As Integer, j As Integer, k As Byte
    
    ' Controlo que no supere los límites
    ' I check that the copy is before the image end
    If xOffset + Width > picWidth Or yOffset + Height > picHeight Then
        ' I return false when a block can't be copied
        CopyBlock = False
        Exit Function
    End If
    
    ' I redim the dest matrix
    ReDim Dest(Width - 1, Height - 1, 2)

    For j = 0 To Height - 1  ' Pos Y
        For k = 0 To 2       ' RGB
            ' Copia los BlockWidth pixels mediante la función copymemory que es mucho más rápida que VB
            ' Copy the entire width size of matrix with CopyMemory Api to increase a lot the speed
            ' It can be done with a simple Dest(x,y,k)=Origen(x + xOffset, j + yOffset,k)
            CopyMemory ByVal VarPtr(Dest(0, j, k)), ByVal VarPtr(Origen(xOffset, j + yOffset, k)), Width
        Next k
    Next j

    ' Return true
    CopyBlock = True
    Exit Function
Solucion:
    If Err.Number = 9 Then
        CopyBlock = False
    Else
        MsgBox "error " & Err.Number & " in copyblock", vbCritical
    End If
End Function

' Escaneo la imagen para ver si hay bloques similares
' This is the main function of the algorithm. It will check if there is a similar block like the compare block
' Parameters:
'       CompareBlock: A matriz like m(BlockWidth,BlockHeight,2--> used for RGB) to be compared
'       iX and iY: Start scanning
'       Igualdad: The percentaje of similar that the blocks should be
'       ClonePos(): The returned data if it function found a similars blocks
Private Function ViewClones(CompareBlock() As Byte, iX As Integer, iY As Integer, Igualdad As Byte, ClonesPos() As tClonePos) As Boolean
    Dim X As Integer, Y As Integer
    Dim IsClone() As Byte
    ReDim ClonesPos(0) As tClonePos
    
    ' Recorro toda la imagen tomando bloques
    ' Run trought the image to find similar blocks
    For Y = iY To picHeight Step BlockHeight
        For X = 0 To picWidth
            
            ' Only for guide uncomment to see the progress of what piece of image it is checking
            'If ShowBlockToVisualHelp = True Then
            '    frmMain.shpClone.Move x, Y
            '    DoEvents
            'End If
            
            ' Si el bloque que voy a analizar (el clon) no se superpone con el Master
            ' If the Block to Check is not overlapped with the master...
            If X >= iX + BlockWidth Or Y <> iY Then
                ' Copio el bloque a analizar (si no lo pude copiar es que el bloque está muy pegado al borde y no alcanzan los pixeles para completar un bloque)
                ' Copy the Block to Check
                If CopyBlock(bPixels(), X, Y, BlockWidth, BlockHeight, IsClone()) = True Then
                    ' Ahora pregunto si el bloque que tome como supuesto master es similar al que tomé ahora
                    ' Now i check the block
                    If IsSimilar(CompareBlock(), IsClone(), Igualdad) = True Then
                        ' Si es asi incremento la matriz de clones...
                        ' If it block is similar than the CompareBlock i save the Clone block position
                        ReDim Preserve ClonesPos(UBound(ClonesPos()) + 1) As tClonePos
                        ' ...guardando su posición en la imagen
                        ClonesPos(UBound(ClonesPos())).X = X
                        ClonesPos(UBound(ClonesPos())).Y = Y
                        ' Notifico que el master tuvo por lo menos una coincidencia
                        ' Notify that the function found at least one block
                        ViewClones = True
                        
                        ' Only for visual guide
                        If ShowBlockToVisualHelp = True Then
                            frmMain.shpCloneLast.Move X, Y
                            DoEvents
                        End If
                        
                        ' Y incremento en BlockWidth pixels el recorrido de clones
                        ' Add a 8 pixels to x coordinate (it prevent an overlapping clone block)
                        X = X + BlockWidth - 1
                    End If
                End If
            End If

        Next X
    Next Y
End Function

' Compara dos bloques si son parecidos
' This function check if sBlock is similar than pBlock
' Parameters:
'       pBlock: The primary matrix (with image block info) to check
'       sBlock: The secondary block
'       Igualdad: The percentage of error
Private Function IsSimilar(pBlock() As Byte, sBlock() As Byte, Igualdad As Byte) As Boolean
    Dim X As Integer, Y As Integer, k As Byte
    Dim Error As Byte
    Dim ByteMayor As Integer
    Dim ByteMenor As Integer
    
    ' Calculo el +/- error de cada pixels
    ' I calculate the tolerance beetween pBlock pixels and sBlock
    Error = 255 - (Igualdad * 255 / 100)
    
    ' Recorro cada pixel en la matriz
    ' Run trought all pixels of the matrix and compare
    For X = 0 To BlockWidth - 1
        For Y = 0 To BlockHeight - 1
            For k = 0 To 2  ' RGB
                
                ' Calculo los valores minimos y máximos de un pixel para saber si es similar o no
                ' Calculate the tolerance of the pixel color that the secondary block should have to be similar
                ByteMayor = CInt(pBlock(X, Y, k)) + Error
                ByteMenor = CInt(pBlock(X, Y, k)) - Error
                ' Arreglo los límites (si es negativo o mayor a 255)
                ' Check for errors
                If ByteMenor < 0 Then ByteMenor = 0
                If ByteMayor > 255 Then ByteMayor = 255

                ' Si el byte(una parte RGB del pixel) no es parecido...
                ' If this pixel isn't similar...
                If ByteMayor < sBlock(X, Y, k) Or ByteMenor > sBlock(X, Y, k) Then
                    ' No es similar y me voy
                    ' ...The blocks aren't similar and go out
                    IsSimilar = False
                    Exit Function
                End If
                
            Next k
        Next Y
    Next X
    ' Si compare todos los pixeles fueron todos similares retorno verdadero
    ' If all pixels were similar return true
    IsSimilar = True
End Function

' Return if a XYPixel is not used for a clone or a master
Private Function NotOverlayed(X As Integer, Y As Integer) As Boolean
    ' Devuelvo el valor negado de la matriz pUsed (que marca los pixels usados)
    NotOverlayed = Not (pUsed(X, Y))
End Function

' Con este procedimiento guardo toda la información recolectada por la función principal en un archivo
' I save the MSF compressed file
' Parameters:
'       Dest: The destination file path
'       MastersBlocks: Master blocks info to be saved
'       CloneBlock: The same but with Clones Blocks
'       Header: The header info to be saved
Private Sub SaveTheFile(Dest As String, MasterBlocks() As tMasterBlock, CloneBlocks() As tCloneBlock, Header As tHeader)
    Dim i As Integer, j As Integer
    Dim PreClonInfo As tPreClone
    
    ' Esta es la estructura del archivo
    ' This is the struct of file
    
    '+---------------------+
    '|     FILE HEADER     |
    '+---------------------+
    '|    MASTER BLOCK 0   |
    '+---------------------+
    '|CANT OF CLONES OF M-0|
    '+---------------------+
    '|   CLONE 0 OF M-0    |
    '+---------------------+
    '|   CLONE 1 OF M-0    |
    '+---------------------+
    '|   CLONE n OF M-0    |
    '+---------------------+
    '|    MASTER BLOCK 1   |
    '+---------------------+
    '|CANT OF CLONES OF M-1|
    '+---------------------+
    '|   CLONE 0 OF M-1    |
    '+---------------------+
    '|   CLONE n OF M-1    |
    '+---------------------+
    '|   FREE PIXEL 0      |
    '+---------------------+
    '|   FREE PIXEL n      |
    '+---------------------+
    '|   END OF FILE (EOF) |    <-- Not used because is not necesary
    '+---------------------+
    
    
    ' Grabo el archivo
    ' Open for save the file
    Open Dest For Binary Access Write As #1
        ' Escribo la info del encabezado
        ' Write the header info
        Put 1, , Header
        
        ' Escribo cada master y cada clon
        ' Write every master and clon
        For j = 1 To UBound(MasterBlocks())
        
            ' Escribo el master n
            ' Write Master n
            Put 1, , MasterBlocks(j)
            
            ' Escribo cuantos clons hay apara este master
            ' I write how much clones there are for this master
            PreClonInfo.Cant = UBound(CloneBlocks(j).Info())
            Put 1, , PreClonInfo.Cant
            
            ' Ahora recorro cada clon y escribo su info
            ' Run trought every clon and write them to file
            For i = 1 To UBound(CloneBlocks(j).Info())
                Put 1, , CloneBlocks(j).Info(i)
            Next i
        Next j
        
        ' Escribo cada pixel libre (sin que este en un clon o en un master)
        ' Write every FreePixels (the pixel who isn't used for a clone or master)
        Dim pPixel As tPixelInfo, Counter As Long
        
        For j = 0 To picHeight - 1
            For i = 0 To picWidth - 1
                ' Controlo que el pixels no haya sido usado
                ' If this pixel is really a free pixel
                If NotOverlayed(i, j) = True Then
                    With pPixel
                        .Red = bPixels(i, j, 2)
                        .Green = bPixels(i, j, 1)
                        .Blue = bPixels(i, j, 0)
                    End With
                    ' Contador de pixels libres
                    ' Count the freepixel (for debug only)
                    Counter = Counter + 1
                    ' Escribo el pixel en el archivo
                    ' Write the freepixel color info
                    Put 1, , pPixel
                End If
            Next i
        Next j
        ' Informo la cantidad de pixels libres
        ' Info the freepixels amount
        MsgBox "Quedaron " & Counter & " pixeles libres" & vbCrLf & _
               "There are " & Counter & " freepixels", vbInformation
        MsgBox "The file were saved", vbInformation
    Close #1
End Sub
