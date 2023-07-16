
      {---------------------------------------------------------------}
      {--------------  El juego de la vida de Conway  ----------------}
      {      Proyecto para GAIA programado por Ricardo Aranguren      }
      {    Compilable con Turbo-Pascal 6.0 de Borland International   }
      {     Necesita los m¢dulos RAT y VGA256 de Ricardo Aranguren    }
      {                 Pamplona, a 3 de Dic de 1997                  }
      {---------------------------------------------------------------}

 {--------------------------------------------------------------------------}
 {     Requerimientos del sistema:                                          }
 {              PC compatible con tarjeta gr fica VGA y rat¢n.              }
 {              Debe ejecutarse en modo MS-DOS.                             }
 {              Ha de estar presente un controlador de rat¢n en memoria.    }
 {--------------------------------------------------------------------------}

PROGRAM CONWAY;

{ Directivas para el compilador: }

{{$DEFINE DEBUG} {A¤ada una "{" al principio antes de compilar un EXE}

{$A+,B-,F-,G+,O-,V-,X+}
{$N-,E-} {Coprocesador no necesario}

{$IFNDEF DEBUG}
  {$D-,R-,S-,I-,L-}
{$ELSE}
  {$D+,R+,S+,I+,L+}
{$ENDIF}


USES rat,vga256,crt;

CONST
  MaxTamMundoX=159;
  MaxTamMundoY=99;
  TamMundoIniX=79;
  TamMundoIniY=49;
  MaxVejez=20;  { Es la m xima edad de vida de una c‚lula representable }
  ProbabilidadMutacion=1;  { Es un tanto por ciento }

TYPE
  TipoCelula= byte; { con valor 0 est  muerta, otro valor indica su edad }
  TipoMundo= array[1..160,1..100] of TipoCelula;

VAR
  TamMundoX, TamMundoY : integer; { Son las dimensiones del mundo}
  Mundo:array[0..1] of TipoMundo;
  NumMundoAct,NumMundoAnt:byte;  { Van cambiando de 0 a 1, y as¡ es m s }
                                 { r pido el cambio de un estado a otro }

{----------------------------------------------------------------------}
{ Esta parte del c¢digo es la dedicada a la interacci¢n con el usuario }
{----------------------------------------------------------------------}

CONST
  TamVentanaX=320;
  TamVentanaY=200;
  IniVentanaX=0;
  IniVentanaY=0;

  ColorCuadr=3;      { Color de la cuadr¡cula }
  ColorCuadrAct=11;  { Color de la cuadr¡cula apuntada por el rat¢n }
  ColorCelulaMuerta=16;

{ Si no controlas el lenguaje ensamblador, te aconsejo que vayas a     }
{ la £ltima parte, la de los algoritmos del juego de la vida de Conway }

VAR
  TamCeldaX, TamCeldaY, IniMundoX, IniMundoY:integer;
  VirtualPant:^TipoPant;

procedure CalculaCuadriculado;
  begin
    TamCeldaX:=pred(TamVentanaX) div TamMundoX;
    TamCeldaY:=pred(TamVentanaY) div TamMundoY;
    if TamCeldaX>TamCeldaY then TamCeldaX:=TamCeldaY;
    if TamCeldaX<TamCeldaY then TamCeldaY:=TamCeldaX;
    IniMundoX:=IniVentanaX+ (TamVentanaX-TamCeldaX*TamMundoX) div 2 ;
    IniMundoY:=IniVentanaY+ (TamVentanaY-TamCeldaY*TamMundoY) div 2 ;
    {El mundo (cuadriculado) est  situado en el centro de la ventana gr fica }
  end;

Procedure DibuCuadricula(x,y,c:byte;var Pant:TipoPant); assembler;
  label LineaHueca;
  { Vamos a dibujar una cuadr¡cula desde el punto                  }
  {    (IniMundoX+TamCeldaX*(x-1),IniMundoY+TamCeldaY*(y-1))       }
  { hasta el punto                                                 }
  {    (IniMundoX+TamCeldaX*x,IniMundoY+TamCeldaY*y)               }
  asm
    dec x   { Lo primero, para simplificar c lculos }
    dec y   { As¡ el rango es de 0 a TamMundoY-1 en vez de 1 a TamMundoY }

    les di,Pant   { ES:DI apunta a Pant }
    mov ax,TamCeldaY
    mul y
    mov bx,320
    mul bx
    add di,ax             { DI= TamCeldaY *y *320 }
    mov ax,TamCeldaX
    mul x
    add di,ax             { DI= TamCeldaY *y *320 + TamCeldaX *x }
    add di,IniMundoX
    mov ax,IniMundoY
    mul bx {320}
    add di,ax             { DI apunta a la esquina izq sup de la cuadr¡cula }

    mov bx,TamCeldaX
    dec bx                { BX almacena el ancho interior de la cuadr¡cula }
    mov dx,319
    sub dx,TamCeldaX      { DX es lo que avanza para empezar la sig. l¡nea }
    mov al,c              { El acumulador guarda el color para pintar }

    mov cx,TamCeldaX
    inc cx
    rep stosb             { L¡nea superior de la cuadr¡cula }
    add DI,DX

    mov CX,TamCeldaY
    dec cx
    LineaHueca:
      stosb
      add DI,BX
      stosb
      add DI,DX
    dec CX
    jnz LineaHueca

    mov cx,TamCeldaX
    inc cx
    rep stosb             { L¡nea inferior de la cuadr¡cula }
  end;

Procedure DibuCelula(x,y,c:byte;Pant:TipoPant); assembler;
  label Linea;
  { Vamos a rellenar la cuadr¡cula anterior                        }
  {    (IniMundoX+TamCeldaX*(x-1),IniMundoY+TamCeldaY*(y-1))       }
  {    (IniMundoX+TamCeldaX*x,IniMundoY+TamCeldaY*y)               }
  asm
    dec x   { Lo primero, para simplificar c lculos }
    dec y   { As¡ el rango es de 0 a TamMundoY-1 en vez de 1 a TamMundoY }

    les di,Pant
    mov ax,TamCeldaY
    mul y
    mov bx,320
    mul bx
    add di,ax             { DI= TamCeldaY *y *320 }
    mov ax,TamCeldaX
    mul x
    add di,ax             { DI= TamCeldaY *y *320 + TamCeldaX *x }
    add di,IniMundoX
    mov ax,IniMundoY
    mul bx {320}
    add di,ax             { DI apunta a la esquina izq sup de la cuadr¡cula }
    add di,321            { DI apunta a la esquina izq sup de la c‚lula }

    mov bl,byte(TamCeldaX)
    dec bl                { BL almacena el ancho de la c‚lula. Recarga CX}
    mov bh,byte(TamCeldaY)
    dec bh                { BH el largo. Actuar  como contador vertical }
    mov dx,321
    sub dx,TamCeldaX      { DX es lo que avanza para empezar la sig. l¡nea }
    mov al,c              { El acumulador guarda el color para pintar }

    xor ch,ch
    Linea:
      mov Cl,bl
      rep stosb
      add DI,DX
    dec bh
    jnz Linea

  end;

Procedure DibuCuadriculado;
  var x,y:integer;
  begin
    for x:=1 to TamMundoX do for y:=1 to TamMundoY do begin
      dibuCuadricula(x,y,ColorCuadr,VirtualPant^);
    end;
  end;

Procedure DibuMundo(var mundoAct:TipoMundo);
  var x,y:integer;
  begin
    for x:=1 to TamMundoX do for y:=1 to TamMundoY do begin
      if Mundo[0,x,y]<>Mundo[1,x,y] then begin
        dibuCelula(x,y,ColorCelulaMuerta+mundoAct[x,y],VirtualPant^);
        Mundo[NumMundoAnt,x,y]:=MundoAct[x,y];
      end;
    end;
    MuestraPant(VirtualPant^);
  end;

VAR UltimoX,UltimoY, PosRatX,PosRatY: integer;

procedure CalculaPosRat; { adapta las coordenads del rat¢n al cuadriculado }
  begin
    UltimoX:=PosRatX;   { Guardamos la posici¢n anterior }
    UltimoY:=PosRatY;
    PosRatX:=1+ (RatX * TamMundoX) div 320;  { Calculamos la nueva }
    PosRatY:=1+ (RatY * TamMundoY) div 200;
    if (UltimoX<>PosRatX) or (UltimoY<>PosRatY) then begin
      { Si el bot¢n izquierdo est  pulsado y cambiamos de casilla, es como }
      { si solt ramos el bot¢n, movieramos a la otra y lo puls ramos. }
      { Esto da facilidad a la hora de editar }
      if ratizq then
        RatIzqClic:=true;  { El m¢dulo RAT permite cambiar esta variable }
      { Debemos dibujar el recuadro del rat¢n s¢lo cuando ‚ste se mueve }
      DibuCuadricula(UltimoX,UltimoY,ColorCuadr,VirtualPant^);
      DibuCuadricula(PosRatX,PosRatY,ColorCuadrAct,VirtualPant^);
    end;
  end;

procedure IniPaletaColores;
  var paleta:TipoPaleta;
      n,f:byte;
  procedure poncolor(c,r,g,b:byte);
    begin
      paleta[c,0]:=r;  { rojo   Red    R }
      paleta[c,1]:=g;  { verde  Green  G }
      paleta[c,2]:=b;  { azul   blue   B }
    end;
   const maxRGB=63;
  begin
    poncolor(0,0,0,0); { el fondo negro }
    poncolor(ColorCuadr,0,20,20);
    poncolor(ColorCuadrAct,0,40,40);
    poncolor(ColorCelulaMuerta,0,0,0);
    for n:=1 to MaxVejez do begin
      f:=50-40*n div MaxVejez;
      poncolor(ColorCelulaMuerta+n,f,50,f);
      { As¡ hacemos que al envejecer la c‚lula cambie de blanco a verde }
    end;
    ponpaleta(paleta);
  end;

procedure IniCuadriculado;
  begin
    BorraPant(VirtualPant^);
    CalculaCuadriculado;
    DibuCuadriculado;
    fillchar(Mundo[NumMundoAnt],MaxTamMundoX*MaxTamMundoY,0);
    DibuMundo(Mundo[NumMundoAct]);
    PosRatX:=1; PosRatY:=1;
  end;

procedure IniInterface;
  begin
    modoGrafico;  { Vga256 }
    new(VirtualPant);
    fillchar(VirtualPant^,SizeOf(TipoPant),0);
    inipaletacolores;
    iniCuadriculado;
    leerat;
    PosRatX:=1;
    PosRatY:=1;
    CalculaPosRat;
  end;

procedure FinInterface;
  begin
    dispose(VirtualPant);
    modoTexto;  { Vga256 }
  end;

{-----------------}
{  INSTRUCCIONES  }
{-----------------}
procedure instrucciones;
  procedure writeCentro(y:integer;cad:string);
    var tam,n:integer;
    begin
      tam:=0;
      for n:=1 to byte(cad[0]) do
        if cad[n]>#31 then inc(tam);
      gotoxy(40-tam div 2,y);
      for n:=1 to byte(cad[0]) do
        if cad[n]>#31 then write(cad[n])
        else textattr:=byte(cad[n]);
    end;
  function encadena(n:integer;ch:char):string;
    var cad:string;
    begin
      cad:='';
      for n:=1 to n do cad:=cad+ch;
      encadena:=cad;
    end;
  begin
    modoTexto;
    clrscr;
    writeCentro(2,#$F+encadena(40,'-'));
    writeCentro(3,#$A'El juego de la vida de CONWAY');
    writeCentro(4,#$F+encadena(40,'-'));
    writeCentro(6,#7'Programado por Ricardo Aranguren (Pamplona, 1997) en colaboraci¢n con '#$A'GAIA'#7',');
    writeCentro(7,'un proyecto de divulgaci¢n de Inteligencia y vida artificial residente en');
    writeCentro(8,#$F'http://www.geocities.com/SiliconValley/Vista/7491'#7);
    writeCentro(9,'(El c¢digo fuente en Turbo-Pascal 6.0 est  disponible)');
    writeCentro(10,'Para cualquier consulta no dude en escribir a: '#2'ra20914@zurron.upna.es');
    writeCentro(12,#$C'---- INSTRUCCIONES DE USO ----'#7);
    writeCentro(13,'Utilice las teclas del cursor para cambiar el tama¤o del juego.');
    writeCentro(14,'Configure las c‚lulas con el bot¢n izquierdo del rat¢n.        ');
    writeCentro(15,'Avance un tiempo con el bot¢n derecho.                         ');
    writeCentro(16,'Sostenga el bot¢n izquierdo y vaya pulsando el derecho para    ');
    writeCentro(17,'    dormir a todas las c‚lulas y despertarlas gradualmente.    ');
    writeCentro(18,'Sostenga el bot¢n derecho y pulse el izquierdo para avanzar    ');
    writeCentro(19,'    en el tiempo autom ticamente.                              ');
    writeCentro(20,'Detenga el proceso autom tico con cualquier bot¢n del rat¢n.   ');
    writeCentro(21,'Desplace todas las c‚lulas usando Supr,Fin,PagAbj,Inicio.      ');
    writeCentro(22,'Controle el modo seguimiento pulsando la barra espaciadora.    ');
    writeCentro(23,'Pulse ESC para salir del juego de la vida.                     ');
    writeCentro(25,#3'Pulsa un bot¢n para continuar');
    repeat leerat until ratIzqClic or ratDerClic or keypressed;
    textattr:=7;

    modoGrafico;  { Vga256 }
    inipaletacolores;
    iniCuadriculado;
  end;

{-----------------------------------------------------------------}
{ Esta parte define los algoritmos del juego de la vida de Conway }
{-----------------------------------------------------------------}

procedure AzarMundo(var QueMundo:TipoMundo; TantoPorCiento:integer);
  var x,y:integer;
  begin
    randomize;
    for x:=1 to TamMundoX do for y:=1 to TamMundoY do
      if random(100)<TantoPorCiento then QueMundo[x,y]:=1
                                    else QueMundo[x,y]:=0;
    DibuMundo(QueMundo);
  end;

procedure IniMundo;
  begin
    fillchar(Mundo,2*MaxTamMundoX*MaxTamMundoY,0); {Pone muertas las c‚lulas}
    TamMundoX:=TamMundoIniX;
    TamMundoY:=TamMundoIniY;
    NumMundoAct:=0;
    NumMundoAnt:=1;
  end;

procedure CambiaEstado(var MundoAnt,MundoAct:TipoMundo);
  var x,y,x1,y1,cont:integer;
  function viva(a,b:integer):boolean;
    begin
      if a>TamMundoX then a:=a-TamMundoX  { No hay c‚lulas fuera del tablero }
      else if a<1 then a:=a+TamMundoX;    { Si se salen por un lado entran }
      if b>TamMundoY then b:=b-TamMundoY  { por el otro }
      else if b<1 then b:=b+TamMundoY;
      viva:=MundoAnt[a,b]>0;
    end;
  begin
    for y:=1 to TamMundoY do for x:=1 to TamMundoX do begin
      cont:=0;
      for y1:=y-1 to y+1 do for x1:=x-1 to x+1 do
        if viva(x1,y1) then inc(cont);
      if viva(x,y) then if cont=4 then cont:=3;
      if cont =3 then begin
        {sigue viva o revive}
        if MundoAnt[x,y]<MaxVejez then inc(MundoAct[x,y]);
      end else MundoAct[x,y]:=0;
    end;
  end;

procedure CambiaAntAct; assembler;
    asm
      xor NumMundoAct,1     { Vamos rotando del mundo 0 al 1 }
      xor NumMundoAnt,1
    end;

procedure finmundo;
  begin
  end;

procedure mutacion;
  begin
    mundo[NummundoAct,1+random(TamMundoX),1+random(TamMundoY)]:=1;
  end;

procedure CambiaCelula(var celula:TipoCelula);
  begin
    if Celula=0 then celula:=1 else celula:=0;
  end;

var editar:boolean;  { Se activa y desactiva con el boton derecho }
                     { Cuando est  activa se puede "pintar", y }
                     { cuando esta inactiva el mundo se mueve }

procedure MueveCelulas(despX,despY:integer);
  var x,y,ox,oy:integer;
  begin
    CambiaAntAct;
    for x:=1 to TamMundoX do begin
      ox:=x+despX;
      if ox>TamMundoX then ox:=ox-tamMundoX
      else if ox<1 then ox:=ox+TamMundoX;
      for y:=1 to TamMundoY do begin
        oy:=y+despY;
        if oy>TamMundoY then oy:=oy-tamMundoY
        else if oy<1 then oy:=oy+TamMundoY;
        Mundo[NumMundoAct,x,y]:=Mundo[NumMundoAnt,ox,oy]
      end;
    end;
  end;

procedure CambiaTamMundo(incX,incY:integer);
  procedure Cambia(var num,inc:integer;Max:integer);
    begin
      num:=num+inc;
      if num<1 then num:=1
      else if num>max then num:=max;
    end;
  begin
    Cambia(TamMundoX,incX,MaxTamMundoX);
    Cambia(TamMundoY,incY,MaxTamMundoY);
  end;

procedure HacerSeguimiento;
  function HayAlgoHoriz(y:integer):boolean;
    var x:integer;
    begin
      x:=TamMundoX;
      while boolean(x) and not boolean(mundo[NumMundoAct,x,y]) do dec(x);
      HayAlgoHoriz:=boolean(x);
    end;
  function HayAlgoVert(x:integer):boolean;
    var y:integer;
    begin
      y:=TamMundoY;
      while boolean(y) and not boolean(mundo[NumMundoAct,x,y]) do dec(y);
      HayAlgoVert:=boolean(y);
    end;
             { Las funciones anteriores devuelven Falso s¢lo si la  }
             { columna/fila indicada contiene s¢lo c‚lulas muertas. }
  begin
    if HayAlgoHoriz(TamMundoY) then CambiaTamMundo(0,1)
    else if not HayAlgoHoriz(pred(TamMundoY)) then CambiaTamMundo(0,-1);
    if HayAlgoHoriz(1) then begin
      CambiaTamMundo(0,1);
      MueveCelulas(0,-1);
    end else if not HayAlgoHoriz(2) then begin
      MueveCelulas(0,1);
      CambiaTamMundo(0,-1);
    end;
    if HayAlgoVert(TamMundoX) then CambiaTamMundo(1,0)
    else if not HayAlgoVert(pred(TamMundoX)) then CambiaTamMundo(-1,0);
    if HayAlgoVert(1) then begin
      CambiaTamMundo(1,0);
      MueveCelulas(-1,0);
    end else if not HayAlgoVert(2) then begin
      MueveCelulas(1,0);
      CambiaTamMundo(-1,0);
    end;
    IniCuadriculado;
  end;

var lleno:integer;
    tecla:char;
    final:boolean;
    seguimiento:boolean;

BEGIN
  randomize;
  IniMundo;
  IniInterface;
  Instrucciones;
  editar:=true;
  final:=false;
  seguimiento:=false;
  repeat
    leerat;
    if keypressed then begin
      while keypressed do tecla:=upcase(readkey); { esto evita los pitidos }
      case tecla of
        #27 {ESC} : final:=true;
        ' ' {esp} : seguimiento:=not seguimiento;
        'K' {Izq} : begin
               CambiaTamMundo(-1,0);
               IniCuadriculado;
           end;
        'M' {Der} : begin
             cambiaTamMundo(1,0);
             IniCuadriculado;
           end;
        'H' {Arr} : begin
             CambiaTamMundo(0,-1);
             IniCuadriculado;
           end;
        'P' {Abj} : begin
             CambiaTamMundo(0,1);
             IniCuadriculado;
           end;
        'S' {Supr}: begin
             MueveCelulas(1,0);
             dibuMundo(Mundo[NumMundoAct]);
           end;
        'Q' {PAbj}: begin
             MueveCelulas(-1,0);
             dibuMundo(Mundo[NumMundoAct]);
           end;
        'G' {Ini}: begin
             MueveCelulas(0,1);
             dibuMundo(Mundo[NumMundoAct]);
           end;
        'O' {Fin}: begin
             MueveCelulas(0,-1);
             dibuMundo(Mundo[NumMundoAct]);
           end;
        else instrucciones;
      end;
    end;
    if seguimiento then HacerSeguimiento;
    if editar then begin
      calculaPosRat;
      if RatIzq and RatDerClic then begin
         lleno:=0;
         while ratizq do begin
           if RatDerClic then begin
             AzarMundo(Mundo[NumMundoAct],lleno);
             lleno:=(lleno+5) mod 101;
           end;
           leerat;
         end;
      end else if ratder and ratizqClic then begin
        editar:=false;
        dibuCuadricula(PosRatX,PosRatY,ColorCuadr,VirtualPant^)
      end else if RatIzqClic then
        CambiaCelula(Mundo[NumMundoAct,PosRatX,PosRatY])
      else if RatDerClic then begin
        CambiaAntAct;
        CambiaEstado(Mundo[NumMundoAnt],Mundo[NumMundoAct]);
      end;
    end else begin
      CambiaAntAct;
      if random(100)<ProbabilidadMutacion then Mutacion;
      CambiaEstado(Mundo[NumMundoAnt],Mundo[NumMundoAct]);
      if RatDerClic or RatIzqClic then begin
        editar:=true;
        dibuCuadricula(PosRatX,PosRatY,ColorCuadrAct,VirtualPant^)
      end;
    end;
    DibuMundo(Mundo[NumMundoAct]);
  until final;
  FinMundo;
  FinInterface;
  if not editar then
    writeln('¨Crees que es ‚ticamente correcto lo que acabas de hacer?');

  { ¨Por qu‚ son cosas distintas matar una persona, un perro, una mosca o   }
  { detener el juego de la vida de Conway? La diferencia radica en el nivel }
  { de complejidad, nada m s. Cuando un sistema artificial sea tan complejo }
  { como nosotros, habr  que replantearse los derechos humanos.             }
  {                                        --- Ricardo Aranguren Gassis --- }
  {                                        ---- ra20914@zurron.upna.es ---- }
END.