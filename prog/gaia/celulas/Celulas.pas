(**************************************************************)
(*           Programa CELULAS  para TURBO PASCAL              *)
(*                                                            *)
(*             (c) Farid Fleifel Tapia, 1995                  *)
(**************************************************************)

(* Controles: Tecla "+": aumenta la velocidad.
              Tecla "-" reduce la velocidad.
              Cualquier otra tecla finaliza el programa. *)

program celulas;
uses crt;

type mundo=array[1..80,1..49] of boolean; (*Matriz de "c‚lulas"*)

var  m:mundo;
     a,b:word;
     i,j,retardo:integer;
     ch:char;
     fin:boolean;

procedure presentaInst;
begin
clrscr;
TextColor(LightBlue);
writeln('                                CELULAS');
writeln('                              -----------');
writeln;
TextColor(Yellow);
writeln('-> ¨C¢mo funciona este programa?');
TextColor(LightGray);
writeln('En cada ciclo se elige aleatoriamente una de las c‚lulas de la matriz.');
writeln('Esa c‚lula muere, dejando un espacio libre.');
writeln;
writeln('Ese espacio es ocupado inmediatamente de la siguiente forma:');
writeln('Se elige a una de las ocho c‚lulas contiguas a ese espacio vac¡o para ');
writeln('reproducirse, y el lugar dejado por la c‚lula muerta lo ocupa una nueva');
writeln('c‚lula, hija de la escogida, y por lo tanto de su misma especie.');
writeln;
writeln('A partir de este comportamiento tan simple podremos observar como el caos ');
writeln('inicial, en el que las c‚lulas de ambas especies se hallan mezcladas, da paso ');
writeln('a una forma de organizaci¢n en la que las c‚lulas de una misma especie forman ');
writeln('amplios grupos que se desplazan, se estiran y se contraen mientras tratan de ');
writeln('sobrevivir.');
writeln;
textcolor(Yellow);
writeln('Controles:  Incremento de velocidad:  Tecla "+"');
writeln('            Disminuci¢n de velocidad: Tecla "-"');
writeln('            Salir del programa:       Cualquier otra tecla.');
writeln;
textcolor(lightgray);
writeln('Pulsa una tecla para continuar');
writeln;
repeat until keypressed;
ch:=readkey;
end;


procedure llenamundo(var m:mundo;var a,b:word);
(* Llena la matriz de c‚lulas con valores aleatorios *)
var i,j:integer;
begin
a:=0;
b:=0;
  for i:=1 to 80 do
    for j:=1 to 49 do
        if random(2)<1 then begin
                            m[i,j]:=false;
                            inc(a)
                            end

                       else begin
                            m[i,j]:=true;
                            inc(b)
                            end;
end;

procedure llenapantalla(var m:mundo);
(*Rellena la pantalla con los valores almacenados en la matriz*)
var i,j:integer;
begin
  gotoxy(1,1);
  for i:=1 to 80 do
      for j:=1 to 49 do
         begin
          if m[i,j] then TextColor(Red)
                    else TextColor(lightblue);
         write('Û')
         end
end;

procedure cambiacelula(var m:mundo; var a,b:word);
(*Elige a un votante al azar y cambia su voto por el de uno de sus vecinos,*)
(*tambien elegido aleatoriamente*)
var i,j,x,y,k:integer;
begin
     i:=random(80)+1;
     j:=random(49)+1;

     k:=random(8);

     case k of
        0:   begin
             x:=i+1;
             y:=j
             end;

        1:   begin
             x:=i+1;
             y:=j+1
             end;

        2:   begin
             x:=i;
             y:=j+1
             end;

        3:   begin
             x:=i-1;
             y:=j+1
             end;

        4:   begin
             x:=i-1;
             y:=j
             end;

        5:   begin
             x:=i-1;
             y:=j-1
             end;

        6:   begin
             x:=i;
             y:=j-1
             end;

        7:   begin
             x:=i+1;
             y:=j-1
             end
     end;

(*Comprobacion de rango. Con esto convertimos al mundo en toroidal*)
(*O sea, con forma de "donut". Las c‚lulas de la primera l¡nea tienen*)
(*como vecinas a las de la £ltima l¡nea y viceversa. Las c‚lulas de la*)
(*primera columna tienen como vecinas a las de la £ltima y viceversa*)

     if x>80 then x:=1
             else
             if x<1 then x:=80;
     if y>49 then y:=1
             else if y<1 then y:=49;


     if (x<1) then clrscr;

     if m[i,j]<>m[x,y] then begin
                       m[i,j]:=m[x,y];
                       gotoxy(i,j);
                       if m[i,j] then begin
                          TextColor(Red);
                          inc(b);
                          dec(a)
                          end

                          else begin
                          TextColor(lightblue);
                          inc(a);
                          dec(b)
                          end;

                       write('Û');
                       end
end;

procedure CambiaRetardo(var retardo:integer; ch:char);

begin
 if (ch='+') and (retardo>0) then retardo:=retardo-1;
 if (ch='-') and (retardo<10) then retardo:=retardo+1;
end;

begin
 presentaInst;
 textmode(c80+font8x8); {80x50}
 randomize;
 llenamundo(m,a,b);
 llenapantalla(m);
 gotoxy(1,50);
 textcolor(lightGray);
 write('C‚lulas Rojas:        C‚lulas Azules:            Retardo:      Cambio:"+"/"-"');
 fin:=false;
 retardo:=1;
 repeat
       for i:=0 to 50 do
         begin
           for j:=0 to 100 do
             cambiacelula(m,a,b);
           delay(retardo)
         end;
       textcolor(lightGray);
       gotoxy(15,50);
       write(b:5);
       gotoxy(38,50);
       write(a:5);
       gotoxy(59,50);
       write(retardo:2);
       if keypressed then
          begin
            ch:=readkey;
            if (ch<>'+') and (ch<>'-') then fin:=true
                         else CambiaRetardo(retardo,ch);
          end;
 until fin;
 textmode(c80); {80x25}
 writeln('C‚lulas. (c) Farid Fleifel Tapia');
 writeln('mailto:fleifel@geocities.com');
 writeln('http://www.geocities.com/SiliconValley/Campus/7808/');
 writeln

end.
