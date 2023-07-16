unit vga256;
{{$DEFINE COPROCESADOR}

{$A+,B-,F-,G+,O-,V-,X+}
          {G+ activa las ordenes del 286 y superiores}
{$IFDEF COPROCESADOR}
  {$N+,E-}
{$ELSE}
  {$N-,E+}   {Coprocesador no utilizado}
{$ENDIF}

{$IFNDEF DEBUG}
  {$D-,R-,S-,I-,L-}
{$ELSE}
  {$D+,R+,S+,I+,L+}
{$ENDIF}
INTERFACE
type
  TipoPant=array[0..199,0..319] of byte;  { Tipo de pantalla 320x200x256 }
  TipoPaleta=array[0..255,0..2] of byte;  { 0=Rojo, 1=Verde, 2=Azul }
var
  realpant:TipoPant absolute $A000:$0000; { Memoria de video }
PROCEDURE AYUDA;
procedure modografico;
procedure modotexto;
procedure copiapant(var pantOrig,pantDest);
procedure muestrapant(var PantVirt);
procedure borrapant(VAR pant);     {Pone a 0 todos los pixels}
procedure ponpaleta(var pal); {indica a la VGA la paleta deseada}
procedure fundido(var paleta:TipoPaleta;brillo:byte); { % de brillo }
procedure negPaleta(var paleta;brillo:byte);
Procedure PonSprite (VAR Sprite,Pant;PosX,PosY:Integer);
Procedure PonSpriteBorde (VAR Sprite,Pant;PosX,PosY:Integer);
Procedure PonSpriteMasc (VAR Sprite,Pant;PosX,PosY:Integer);
Procedure PonSpriteMascBorde (VAR Sprite,Pant;PosX,PosY:Integer);

IMPLEMENTATION
procedure modografico; assembler;
  asm
    mov aX,0013h    {rutina 0=modo de video, modo 13h=320x200x256}
    int 10h         {llamada a BIOS rutinas de video}
  end;
procedure modotexto; assembler;
  asm
    mov AX,0003h    {rutina 0=modo de video, modo 3h=texto 80x25 color}
    int 10h         {llamada a BIOS rutinas de video}
  end;
procedure copiapant(var pantOrig,pantDest); assembler;
  asm
    PUSH DS
    PUSH ES                   {salva segmentos DS y ES en la pila}
    LDS SI,Pantorig           {carga la direcci¢n de PantOrig. DS=seg SI=ofs }
    LES DI,PantDest           {carga la direcci¢n de PantDest. ES=seg DI=ofs }
    mov cx,32000d             {indicamos el n£mero de repeticiones 64000/2}
    cld                       {Incrementar  DI y SI cada vez}
    rep movsw                 {y el bucle se repite hasta que cx es 0}
    POP ES
    POP DS                    {recuperamos ES y DS de la pila}
  end;
procedure muestrapant(var PantVirt); assembler;
  label B1,B2;
  asm
    PUSH DS                  {guardamos en la pila DS y ES}
    PUSH ES
    LDS SI,PantVirt          {Cargamos la direcci¢n de PantVirt como DS:SI }
    MOV BX,$A000
    MOV ES,BX                {damos a ES el valor $A000 (segmento video VGA)}
    XOR DI,DI                {ES:DI apunta al pixel arriba izquierda}
    MOV CX,32000d            {Cargamos el contador a 64000/2 palabras}
    cld                      {Direcci¢n ascendente (incrementar  ES y DI)}

    MOV DX,03DAh             {Carga en DX el n£mero de puerto de la VGA}
    B1:
      IN AL,DX                 {Leemos del puerto}
      AND AL,8h                {Comprobamos el bit del retrazado}
      JNZ B1                   {Si est  activado, salta a B1}
    B2:
      IN AL,DX                 {Leemos el puerto}
      AND AL,8h                {bit de retrazado}
      JZ B2                    {Si est  activado, contin£a}

    rep movsw                {Volcado de pantalla}
    POP ES
    POP DS                   {recuperamos de la pila ES y DS}
  end;

procedure borrapant(VAR pant); assembler;    {Pone a 0 todos los pixels}
  asm
    PUSH ES           { salvamos el ES en la pila }
    LES DI,Pant       { direcci¢n de Pant= ES:DI }
    mov cx,32000d     { contador a 64000/2 }
    xor ax,ax         { el valor de STOrage es 0 }
    cld               { direcci¢n ascendente (Clear Direction Flag) }
    rep stosw         { REPit STOrage String Word (rellenado a 0) }
    POP ES            { recuperamos ES de la pila }
  end;

procedure ponpaleta(var pal);assembler; {indica a la VGA la paleta deseada}
  LABEL B1,B2;
  asm
    PUSH DS
    LDS SI,pal
    xor al,al            {el primer color que introduciremos}
    MOV DX,$3C8          {puerto $3C8: VGA video DAC PEL address}
    OUT DX,AL            {primer color 0}
    MOV CX,768d          {CX contador de los 256 colores, CX=256*3/2 palabras}
    CLD

    MOV DX,03DAh             {Carga en DX el n£mero de puerto de la VGA}
    B1:
      IN AL,DX                 {Leemos del puerto}
      AND AL,8h                {Comprobamos el bit del retrazado}
      JNZ B1                   {Si est  activado, salta a B1}
    B2:
      IN AL,DX                 {Leemos el puerto}
      AND AL,8h                {bit de retrazado}
      JZ B2                    {Si est  activado, contin£a}

    MOV DX,$3C9          {puerto $3C9: VGA video DAC PEL }
    REP OUTSB            {esto no va a funcionar en un 8086 :(  }
    POP DS
  end;
procedure Fundido(var paleta:TipoPaleta;brillo:byte); { % de brillo }
  var
    x,c:byte;
    palaux:^TipoPaleta;
  begin
    new(palaux);
    if brillo=0 then fillchar(palaux^,256*3,0) else
    for x:=0 to 255 do for c:=0 to 2 do
      palaux^[x,c]:=paleta[x,c]*brillo DIV 100;
    ponpaleta(palaux^);
    dispose(palaux);
  end;
procedure negPaleta(var paleta;brillo:byte); assembler;
  LABEL B1,B2,Pal;
  asm
    PUSH DS
    LDS SI,paleta
    xor al,al            {el primer color que introduciremos}
    MOV DX,$3C8          {puerto $3C8: VGA video DAC PEL address}
    OUT DX,AL            {primer color 0}
    MOV CX,768d          {CX contador de los 256 colores, CX=256*3/2 palabras}
    CLD
    MOV BL,brillo
    MOV BH,100

    MOV DX,03DAh             {Carga en DX el n£mero de puerto de la VGA}
   B1:
    IN AL,DX                 {Leemos del puerto}
    AND AL,8h                {Comprobamos el bit del retrazado}
    JNZ B1                   {Si est  activado, salta a B1}
   B2:
    IN AL,DX                 {Leemos el puerto}
    AND AL,8h                {bit de retrazado}
    JZ B2                    {Si est  activado, contin£a}

    MOV DX,$3C9          {puerto $3C9: VGA video DAC PEL }
    pal:
      LODSB              {carga el valor apuntado por DS:SI en AL}
      MUL BL             {Se multiplica el valor cargado por brillo}
      DIV BH             {el resultado se divide entre 100}
      OUT DX,AL          {y se manda al puerto $3C9}
      LOOP pal            {se repite 256*3 veces}
    POP DS
  end;
Procedure PonSprite (VAR Sprite,Pant;PosX,PosY:Integer); assembler;
{      (0,0)
         X--------------------------------+
         |   (PosX,PosY)                  |
         |        X---------+             |
         |        |         |             |
         |        | Sprite  |TamY         |
         |        |         |             |
         |        |         |             |
         |        +---------+             |
         |           TamX         Pant    |
         |                                |
         +--------------------------------X
                                      (319,199)
}
  label bucley;
  var TamX:word;
  asm
    PUSH DS
    PUSH ES
    LDS SI,Sprite
    LES DI,Pant
    MOV AX,PosY
    MOV CX,320d
    IMUL CX             {AX le damos PosY*320, acarreo en DX se ignora}
    ADD DI,AX          {y vamos a la linea PosY de Pant, DI:=DI+PosY*320}
    ADD DI,PosX        {Posici¢nX en Pant aumenta PosX en DI}
    CLD
    LODSW
    MOV BX,AX            {BX contiene TamX}
    LODSW
    MOV DX,AX            {DX contiene TamY}
    MOV AX,BX          {LongitudX en AX}
    SUB AX,319         {AX contiene -(320-LongitudX)}
    INC DX
    INC BX
    {Tenemos:  DS:SI en comienzo de la primera linea del bitmap a volcar,
               ES:DI en pant[PosY,PosX],
               BX=LongitudX
              -AX=320-LongitudX, lo que aumenta DI cuando se acaba una linea
               DX=LongitudY, s¢lo se usa una vez, y es el contador de lineas
               CX ser  el contador para MOVS, recargado con BX }

    BucleY:         {regreso aqu¡ cada vez que acaba una linea}
      MOV CX,BX     {Recargamos el contador}
      REP MOVSB     {Vuelca la linea... (LongitudX bytes)}
      SUB DI,AX      {DI al principio de la linea siguiente en Pant}
      DEC DL
      JNZ BucleY    {Pone lineas hasta que DL es 0 tras decrementarlo}
    POP ES
    POP DS
  end;
Procedure PonSpriteBorde (VAR Sprite,Pant;PosX,PosY:Integer); assembler;
{      (0,0)
         X--------------------------------+
         |   (PosX,PosY)                  |
         |        X---------+             |
         |        |         |             |
         |        | Sprite  |TamY         |
         |        |         |             |
         |        |         |             |
         |        +---------+             |
         |           TamX         Pant    |
         |                                |
         +--------------------------------X
                                      (319,199)
}
  label final,fuerafiny,calcpy,fuerapy,calcx,dentrofinx,dentroposx,dibujar,bucley;
  var TamX:word;
  asm
    PUSH DS
    PUSH ES
    LDS SI,Sprite
    LES DI,Pant
    CLD
    LODSW
    MOV TamX,AX
    LODSW                {AX contiene TamY}
    ADD AX,PosY          {AX contiene FinY=PosY+TamY}
    CMP AX,0
    JL final             {Si FinY<0 no se dibuja}
    CMP AX,199d
    JG FueraFinY
        MOV BX,AX      {Si FinY<200 , BetaY=FinY, en BX}
        JMP CalcPy
      FueraFinY:       {Si FinY>199 , BetaY=199, en BX}
        MOV BX,199d
    CalcPy:
    MOV AX,PosY
    CMP AX,200d
    JGE Final          {Si PosYò200 , no se dibuja}
    CMP AX,0
    JL FueraPy         {Si PosY<0 , se sale por arriba}
        SUB BX,AX          {LongitudY=BetaY-PosY en BX}
                           {SI se queda igual, en la primera linea del sprite}
        MOV CX,320d
        IMUL CX             {AX le damos PosY*320, acarreo en DX se ignora}
        ADD DI,AX          {y vamos a la linea PosY de Pant}
        JMP CalcX
      FueraPy:
                           {LongitudY=BetaY en BX}
        MOV DX,TamX
        INC DX
        IMUL DX           {AX contiene PosY*(TamX+1), acarreo en DX lo ignoro}
        SUB SI,AX          {Vamos a la linea PosY del sprite}
                           {DI apunta a la linea Y=0 de Pant}
    CalcX:
    MOV DX,TamX          { DX contiene TamX }
    MOV CX,PosX          { CX contiene PosX }
    MOV AX,CX
    ADD AX,DX            {AX contiene FinX=PosX+TamX}
    CMP AX,0
    JL Final             {si FinX<0 , no se dibuja}
    CMP AX,319d
    JNG DentroFinX       {Si FinXó319, FinX entra en Pant}
      MOV AX,319d        {Si se corta por la derecha, BetaX=319 en AX}
    DentroFinX:          {si entra, BetaX=FinX en AX}
    CMP CX,320d
    JGE Final          { Si PosXò320 , no dibuja }
    CMP CX,0
    JGE DentroPosX   {Si PosXò0, El Sprite cabe por la Izquierda}
                       {Si se corta el sprite por la izquierda...}
                       {LongitudX=BetaX, la dist hasta el borde derecho en AX}
      SUB SI,CX        {Posici¢nX en el Sprite aumenta PosX}
                       {Posici¢nX en Pant es 0, la actual}
      JMP Dibujar
    DentroPosX:      {si no se corta por la izq...}
                       {Posici¢nX en el Sprite es 0, SI no cambia}
      SUB AX,CX        {LongitudX=BetaX-PosX en AX}
      ADD DI,CX        {Posici¢nX en Pant aumenta PosX en DI}
    Dibujar:
    SUB DX,AX
    INC AX
    INC BL
    {Tenemos:  DS:SI en comienzo de la primera linea del bitmap a volcar,
               ES:DI en pant[PosY,PosX],
               AX=LongitudX
               DX=TamX-AX, lo que aumenta SI cuando se hace una linea
               BL=LongitudY, s¢lo se usa una vez, y es el contador de lineas
               CX ser  el contador para MOVS, recargado cada linea con AX
                  ­­PERFECTO!! (Si funciona me rapo el pelo al 3) :)   }


    BucleY:         {regreso aqu¡ cada vez que acaba una linea}
      MOV CX,AX     {Recargamos el contador}
      REP MOVSB     {Vuelca la linea... (LongitudX bytes)}
      ADD SI,DX     {SI a la linea siguiente}
      ADD DI,320     {DI al final de la linea siguiente Pant}
      SUB DI,AX      {DI al principio de esta linea en Pant}
      DEC BL
      JNZ BucleY    {Pone lineas hasta que BL es 0 tras decrementarlo}
    final:
    POP ES
    POP DS
  end;
Procedure PonSpriteMasc (VAR Sprite,Pant;PosX,PosY:Integer); assembler;
{      (0,0)
         X--------------------------------+
         |   (PosX,PosY)                  |
         |        X---------+             |
         |        |         |             |
         |        | Sprite  |TamY         |
         |        |         |             |
         |        |         |             |
         |        +---------+             |
         |           TamX         Pant    |
         |                                |
         +--------------------------------X
                                      (319,199)
}
  label bucley,buclex,transparente;
  var TamX:word;
  asm
    PUSH DS
    PUSH ES
    LDS SI,Sprite
    LES DI,Pant
    MOV AX,PosY
    MOV CX,320d
    IMUL CX             {AX le damos PosY*320, acarreo en DX se ignora}
    ADD DI,AX          {y vamos a la linea PosY de Pant, DI:=DI+PosY*320}
    ADD DI,PosX        {Posici¢nX en Pant aumenta PosX en DI}
    CLD
    LODSW
    MOV BX,AX            {BX contiene TamX}
    LODSW
    MOV DX,AX            {DX contiene TamY}
    PUSH BP    {cuesta menos acceder a un registro que a memoria}
    MOV BP,BX          {LongitudX en BP}
    SUB BP,319         {BP contiene -(320-LongitudX)}
    INC DX
    INC BX
    {Tenemos:  DS:SI en comienzo de la primera linea del bitmap a volcar,
               ES:DI en pant[PosY,PosX],
               BX=LongitudX
              -BP=320-LongitudX, lo que aumenta DI cuando se acaba una linea
               DX=LongitudY, s¢lo se usa una vez, y es el contador de lineas
               CX ser  el contador para LODS, recargado con BX
               AL es para comprobar si el color es transparente=0 }

    BucleY:         {regreso aqu¡ cada vez que acaba una linea}
      MOV CX,BX     {Recargamos el contador}
      BucleX:
        LODSB
        CMP AL,0
        JE Transparente
        MOV ES:[DI],AL
        Transparente:
        INC DI
        LOOP BucleX
      SUB DI,BP      {DI al principio de la linea siguiente en Pant}
      DEC DL
      JNZ BucleY    {Pone lineas hasta que CX es 0 tras decrementarlo}
    POP BP
    POP ES
    POP DS
  end;
Procedure PonSpriteMascBorde (VAR Sprite,Pant;PosX,PosY:Integer); assembler;
{      (0,0)
         X--------------------------------+
         |   (PosX,PosY)                  |
         |        X---------+             |
         |        |         |             |
         |        | Sprite  |TamY         |
         |        |         |             |
         |        |         |             |
         |        +---------+             |
         |           TamX         Pant    |
         |                                |
         +--------------------------------X
                                      (319,199)
}
  label final,fuerafiny,calcpy,fuerapy,calcx,dentrofinx,dentroposx,dibujar,bucley,buclex,transparente,finalBP;
  var TamX:word;
  asm
    PUSH DS
    PUSH ES
    LDS SI,Sprite
    LES DI,Pant
    CLD
    LODSW
    MOV TamX,AX
    LODSW                {AX contiene TamY}
    ADD AX,PosY          {AX contiene FinY=PosY+TamY}
    CMP AX,0
    JL final             {Si FinY<0 no se dibuja}
    CMP AX,199d
    JG FueraFinY
        MOV BX,AX      {Si FinY<200 , BetaY=FinY, en BX}
        JMP CalcPy
      FueraFinY:       {Si FinY>199 , BetaY=199, en BX}
        MOV BX,199d
    CalcPy:
    MOV AX,PosY
    CMP AX,200d
    JGE Final          {Si PosYò200 , no se dibuja}
    CMP AX,0
    JL FueraPy         {Si PosY<0 , se sale por arriba}
        SUB BX,AX          {LongitudY=BetaY-PosY en BX}
                           {SI se queda igual, en la primera linea del sprite}
        MOV CX,320d
        IMUL CX             {AX le damos PosY*320, acarreo en DX se ignora}
        ADD DI,AX          {y vamos a la linea PosY de Pant}
        JMP CalcX
      FueraPy:
                           {LongitudY=BetaY en BX}
        MOV DX,TamX
        INC DX
        IMUL DX           {AX contiene PosY*(TamX+1), acarreo en DX lo ignoro}
        SUB SI,AX          {Vamos a la linea PosY del sprite}
                           {DI apunta a la linea Y=0 de Pant}
    CalcX:
    MOV DX,TamX          { DX contiene TamX }
    MOV CX,PosX          { CX contiene PosX }
    PUSH BP              {Ya no vamos a cargar m s variables, no necesitamos
                         el puntero base y lo usamos como registro de 16 bits}
    MOV BP,CX
    ADD BP,DX            {BP contiene FinX=PosX+TamX}
    CMP BP,0
    JL FinalBP             {si FinX<0 , no se dibuja}
    CMP BP,319d
    JNG DentroFinX       {Si FinXó319, FinX entra en Pant}
      MOV BP,319d        {Si se corta por la derecha, BetaX=319 en BP}
    DentroFinX:          {si entra, BetaX=FinX en AX}
    CMP CX,320d
    JGE FinalBP          { Si PosXò320 , no dibuja }
    CMP CX,0
    JGE DentroPosX   {Si PosXò0, El Sprite cabe por la Izquierda}
                       {Si se corta el sprite por la izquierda...}
                       {LongitudX=BetaX, la dist hasta el borde derecho en BP}
      SUB SI,CX        {Posici¢nX en el Sprite aumenta PosX}
                       {Posici¢nX en Pant es 0, la actual}
      JMP Dibujar
    DentroPosX:      {si no se corta por la izq...}
                       {Posici¢nX en el Sprite es 0, SI no cambia}
      SUB BP,CX        {LongitudX=BetaX-PosX en BP}
      ADD DI,CX        {Posici¢nX en Pant aumenta PosX en DI}
    Dibujar:
    SUB DX,BP
    INC BP
    INC BL
    {Tenemos:  DS:SI en comienzo de la primera linea del bitmap a volcar,
               ES:DI en pant[PosY,PosX],
               BP=LongitudX
               DX=TamX-AX, lo que aumenta SI cuando se hace una linea
               BL=LongitudY, s¢lo se usa una vez, y es el contador de lineas
               CX ser  el contador para MOVS, recargado cada linea con BP
               y AL servir  para comprobar si el color es transparente=0}

    BucleY:         {regreso aqu¡ cada vez que acaba una linea}
      MOV CX,BP     {Recargamos el contador}
      BucleX:
        MOV AL,[SI]        {el color que toca poner en Pant, cargado en AX}
        AND AL,AL          {Si el color es transparente, AND 0,0 activa FZ}
        JZ Transparente
        MOV ES:[DI],AL
        Transparente:
        INC SI
        INC DI
        LOOP BucleX
      ADD SI,DX     {SI a la linea siguiente}
      ADD DI,320     {DI al final de la linea siguiente Pant}
      SUB DI,BP      {DI al principio de esta linea en Pant}
      DEC BL
      JNZ BucleY    {Pone lineas hasta que CX es 0 tras decrementarlo}
    finalBP:
    POP BP
    final:
    POP ES
    POP DS
  end;

procedure ayuda;
  begin
    writeln('       ---- Unidad VGA256 por Ricardo Aranguren para Turbo Pascal ----      1/2');
    writeln;
    writeln('    En este modulo estan incluidos los procedimientos b sicos para el manejo de');
    writeln('la VGA en modo 320x200x256.  Todas las rutinas fueron escritas en ensamblador y');
    writeln('trabajar n con un 286 o superior.  Los procedimientos son:');
    writeln('    Procedure ModoGrafico');
    writeln('        Inicia el modo de video 320x200x256.(no comprueba nada, s¢lo lo inicia)');
    writeln('    Procedure ModoTexto');
    writeln('        Regresa al modo texto 80x25 a color');
    writeln('    Procedure CopiaPant(var PantOrig,PantDest)');
    writeln('        Copia 64000 bytes desde PantOrig hasta PantDest');
    writeln('    Procedure MuestraPant(var Pant)');
    writeln('        Copia 64000 bytes de Pant al Segmento de video controlando el retrazado');
    writeln('    Procedure BorraPant(var Pant)');
    writeln('        Pone a 0 los 64000 bytes de Pant');
    writeln('    Procedure PonPaleta(var Paleta)');
    writeln('        Manda 768 bytes (256*3) de Paleta a la VGA, controlando el retrazado');
    writeln('    Procedure NegPaleta(var Paleta;TantoPorCien:byte)');
    writeln('        Igual que PonPaleta pero enviando solo el TantoPorCien de cada color');
    writeln('Adem s, est n definidos:');
    write('  TYPE  TipoPant=array[0..199,0..319] of byte;  { Tipo de pantalla 320x200x256 }');
    writeln('        TipoPaleta=array[0..255,0..2] of byte;  { 0=Rojo, 1=Verde, 2=Azul }');
    writeln('  VAR   realpant:TipoPant absolute $A000:0;');
    writeln;
    write('Pulse INTRO...');
    readln;
    writeln('       ---- Unidad VGA256 por Ricardo Aranguren para Turbo Pascal ----      2/2');
    writeln;
    writeln('    Adem s se incluyen procedimientos para manejo de sprites, que deber n tener');
    writeln('el formato siguiente: 2 word de tama¤o X,Y y luego el bitmap, un byte por pixel');
    writeln('as¡ que el tama¤o de un sprite 320x200 ser  4+320*200=64004 bytes.');
    writeln;
    writeln('Los procedimientos incluidos son:');
    writeln('    Procedure PonSprite (VAR Sprite,Pant;PosX,PosY:Integer);');
    writeln('    Procedure PonSpriteBorde (VAR Sprite,Pant;PosX,PosY:Integer);');
    writeln('    Procedure PonSpriteMasc (VAR Sprite,Pant;PosX,PosY:Integer);');
    writeln('    Procedure PonSpriteMascBorde (VAR Sprite,Pant;PosX,PosY:Integer);');
    writeln;
    writeln('  Sirven para volcar en la pantalla Pant el contenido de Sprite, de las cuatro');
    writeln('formas anteriores, expuestas de mayor a menor rapidez de volcado.');
    write  ('  Las coordenadas PosX,PosY son las coordenadas en la pantalla Pant de la esqui-');
    writeln('na superior izquierda de la imagen que queremos mostrar');
    writeln('  Si el procedimiento comprueba el Borde, entonces podemos poner un sprite que');
    writeln('se sale de la pantalla o que es m s grande que ‚sta, o que est  fuera, sin que');
    writeln('aparezca por el otro lado de la pantalla');
    writeln('  Si el procedimiento tiene MASCara, entonces se vuelca el sprite de forma que');
    writeln('cuando encuentre el byte 0, deje ese pixel como est  en la pantalla.');
    writeln('  Recuerda:  El color 0 se utiliza como color transparente');
    writeln;
    writeln('                                                              (Pamplona, 1997)');
    write('Pulse INTRO...');
    readln;
{    writeln('');
    writeln('');
    writeln('');
    writeln('');
    writeln('');
    writeln('');
    writeln('');}
  end;
Function ExisteVGA:boolean;
  var
    dev:byte;
  begin
    asm
      MOV AX,01A00h
      INT 10h
      MOV dev,bl
    end;
    ExisteVga:=(dev=7) or (dev=8);
  end;
begin
  if not existeVGA then begin
    writeln('ERROR: Este equipo no dispone de tarjeta gr fica VGA compatible');
    write('Pulse INTRO para continuar...');
    readln;
    halt(1);
  end;
end.