unit rat;
{ $D-,L-}   { Para evitar el trazado dentro del m¢dulo RAT }
interface
var
	ratx,raty:word;
	ratizq,ratder,ratizqclic,ratderclic:boolean;
procedure inirat;
procedure leerat;
procedure Ayuda;
procedure Punteroon;
procedure punterooff;
procedure punteroaXY(x,y:word);
procedure test;
implementation
procedure ayuda;
  begin
    writeln;
    writeln('             --------  RAT v3.0  por Ricardo Aranguren  --------');
    writeln;
    writeln('  M¢dulo para controlar el rat¢n');
    writeln('Variables accesibles:');
    writeln('  RatX, RatY : word;               {Coordenadas del puntero en ventana 320x200}');
    writeln('  RatIzq,RatDer,                   {boton Izquierdo y Derecho pulsados}');
    writeln('  RatIzqClic,RatDerClic: boolean;  {Botones reci‚n pulsados}');
    writeln('Procedimientos :');
    writeln('  Test : Realiza un test del rat¢n');
    writeln('  IniRat : Inicializa el controlador seg£n el modo gr fico actual');
    writeln('  LeeRat : Actualiza las variables antes indicadas seg£n el estado del rat¢n');
    writeln('  PunteroON,PunteroOFF : Visualizar puntero o no hacerlo, he ah¡ la cuesti¢n');
    writeln('  PunteroaXY(x,y) : Poner el puntero en la posici¢n (x,y)');
    writeln;
    writeln('  Pulse INTRO...');
    readln;
  end;
var
	antizq,antder:boolean;
procedure inirat; assembler;
  asm
    { INT 33,0 }
      MOV AX,0
      mov ratizq,AL   {valores iniciales a FALSE}
      mov ratder,AL
      INT 33h
  end;
procedure PunteroOn; assembler;
  asm
    { INT 33,1 }
      MOV AX,1
      INT 33h
  end;
procedure PunteroOff; assembler;
  asm
    { INT 33,2 }
      MOV AX,2
      INT 33h
  end;
procedure leerat;
  begin
    AntIzq:=RatIzq;
    AntDer:=RatDer;
    asm
      { INT 33,3 para recoger en BX,CX,DX los datos del raton }
      MOV AX,3
      INT 33h
      SHR CX,1        {Dividimos la posici¢n entre dos, para el modo 320x200}
      MOV RatX,CX
      MOV RatY,DX
      MOV DX,BX        {Status del Raton duplicado en BX y en DX}
      AND BX,1         {M scara con Bit 0 del bot¢n izquierdo}
      MOV RatIzq,BL    {Y actualizamos el booleano}
      AND DX,2         {lo mismo con el Bit 1 del bot¢n Derecho}
      MOV RatDer,DL
    end;
    RatIzqClic:= RatIzq and not AntIzq;
    RatDerClic:= Ratder and not Antder;
  end;
procedure punteroaxy(x,y:word); assembler;
  asm
    MOV AX,4
    MOV CX,x
    SHL CX,1
    MOV DX,y
    INT 33h
  end;
function ratoninstalado:boolean;
	var
		intraton:^byte absolute $0000:$00cc;
    ri:boolean;
	begin
		if (intraton<>nil) and (intraton^<>$cf) then
      asm
        MOV AX,0
        INT 33h
        MOV ri,AL
      end else ri:=false;
    ratoninstalado:=ri;
	end;
procedure test;
  begin
    repeat
      leerat;
      writeln(ratx:9,raty:9,ratizq:9,ratder:9,ratizqclic:9,ratderclic:9);
    until ratizqclic and ratderclic;
  end;
begin
	if not ratoninstalado then begin
			writeln('ERROR: controlador de rat¢n no instalado en memoria');
      writeln('Si dispone de un rat¢n, deber¡a tener un disquete con el "driver".');
      writeln('Bajo Windows tambi‚n deber  instalar el controlador.');
      writeln('Consulte la documentaci¢n o pregunte a su servicio t‚cnico.');
      writeln('Pulse INTRO para continuar...');
      readln;
			halt(1);
		end;
  inirat;
	leerat;
end.