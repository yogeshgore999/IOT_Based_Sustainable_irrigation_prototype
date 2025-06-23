Assembly code
org 0000h ;originate from 0000h
ljmp main1 ;then jumps to routine mainl
org 0023h ;serial interupt service routine
Ijmp serial ;jump to serial
main1:
mov p1,#0fth ;input port(n.a.)
mov p0,#00h ;output port(motors)
mov p2,#0ffh ;input port(adc)
mov p3,#0ffh input(n.a)
mov sp,#50h ;stack pointer 50h intialise (ram)
mov r5,#01h ;r5 pc to uc serial data (motor data)
mov r2,#01h
clr p0.4
setb p0.5
ret
s12: clr p0.4
clr p0.5
ret
adel:
clr p0.1 ;clears the write pin
Icall delay0
setb p0.1 ;sets the write pin
Icall delay0 ;waits for 45 msecs and then go to next step
mov a.#’t’
Icall send
Icall delay0
mov a,p1
Icall send
Icall delay0
ret
adc2:
clr p0.1 ;clears the write pin
Icall delay0
setb p0.1 ;sets the write pin
Icall delay0 ;waits for 45 msecs and then go to next step
mov a,#’h’
Icall send
Icall delay0
mov a,p2
Icall send 
Icall delay0 
ret

Delay0 :45 ms delay
mov r3,#255
here2:
mov r4,#255
here1:
djnz r4.here1
djnz r3.here2
ret
mov tmod,#20h; timer1 baud rate set
mov th1,#0fdh:9600
mov scon.#50h;n,1,8
setb TR1 ;start timer
ret
send: mov sbuf,a ;to transmit the data serially
tx: jnb T1, tx ;wait untill t1 is set
clr T1 ;will clear ti so as to transmit next data
ret
delay: ;timer 0 used for delay
mov tmod,#21h
mov rl,#7
back3: mov t10,#00h
mov th0,#00h
setb tr0
again3: jnb tf0,again3
clr tr0
clr rf0
djnz r1,back3
ret
delay1:
mov tmod,#21h
mov r1#14
back: mov t10,#00h
mov th0,#00h
seth tr0
again: jnb tf0, again
clr tr0
clr tf0
djnz r1 back
ret
End
