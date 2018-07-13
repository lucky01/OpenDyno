Dim bttimeout As Word
bttimeout = 0

Dim pulsedisp As Bit
Dim diagdisp As Word

pulsedisp = False
diagdisp = 0

Dim rpm As Word
Dim rpmperiod As Word
Dim rpmout As Word
Dim rpmhighwait As Word
Dim rpmlow As Word

rpm = 0
rpmperiod = 0
rpmout = 0
rpmhighwait = 0
rpmlow = 0

Dim pulseperiod As Word
Dim pulseout As Word
Dim pulsehighwait As Bit
Dim pulselow As Word

pulseout = 0
pulseperiod = 0
pulsehighwait = 0
pulselow = 0

Dim charinput As Byte
charinput = 0
Dim mode As Byte
mode = 0

ConfigPin PORTD.2 = Input  'pulse
ConfigPin PORTD.3 = Input  'spark
ConfigPin PORTD.4 = Output  'LED
ConfigPin PORTB.0 = Input  'reference
ConfigPin PORTB.1 = Input  'lambda input
ConfigPin PORTB.2 = Output  'load Capacitor
ConfigPin PORTB.4 = Input  'button 1
ConfigPin PORTB.5 = Input  'button 2
ConfigPin PORTB.6 = Input  'button 3

PORTD.2 = 0
PORTD.3 = 0
PORTB.0 = 0
PORTB.1 = 0
PORTB.4 = 0
PORTB.5 = 0
PORTB.6 = 0

Symbol bt1 = PINB.4
Symbol bt2 = PINB.5
Symbol bt3 = PINB.6
Symbol pulse = PIND.2
Symbol led = PORTD.4
Symbol spark = PIND.3
Symbol lambda = ACSR.ACO

ACSR.ACD = 0
Disable ACI

'set timer parameters
TIMSK.OCIE1A = 1
TCCR1B.CS12 = 0
TCCR1B.CS11 = 0
TCCR1B.CS10 = 1
TCCR1B.CTC1 = 1
OCR1AH = $00
OCR1AL = $63
Enable OC1A
Enable


Dim getlambda As Bit
Dim lambdaperiod As Word
Dim lambdawait As Word
Dim lambdaout As Word

getlambda = False
lambdaperiod = 0
lambdawait = 0
lambdaout = 0

Hseropen 57600
led = True

main_loop:

Hserget charinput

Select Case charinput
	Case "0"
		mode = 0
		Hserout "STOP", Lf
		led = True
	Case "1"
		mode = 1
		Hserout "DATA", Lf
		led = False
	Case "2"
		mode = 2
		Hserout "DIAG", Lf
		led = False
EndSelect

If bttimeout = 0 Then
If bt1 = 0 Then
	Hserout "BT1", Lf
Endif
If bt2 = 0 Then
	Hserout "BT2", Lf
Endif
If bt3 = 0 Then
	Hserout "BT3", Lf
Endif
bttimeout = 5000
Else
bttimeout = bttimeout - 1
Endif


If pulselow >= 60000 Then
	pulselow = 60000
	pulseperiod = 0
	pulseout = 0
Endif

If rpmlow >= 60000 Then
	rpmout = 0
	rpmperiod = 0
	rpmlow = 60000
Endif

If lambdaperiod >= 50000 Then
	lambdaperiod = 50000
	lambdaout = 0
Endif

Select Case mode
	Case 1
			If pulsedisp = True And pulseperiod < 60000 Then
					Toggle led
					If rpmperiod > 240 And rpmperiod < 12000 Then
						rpmout = rpmperiod
					Endif
					Hserout #pulseperiod, ";", #rpmout, ";", #lambdaout, ";", Lf
					pulsedisp = False
			Endif
	Case 2
			If diagdisp >= 1000 Then
					diagdisp = 0
					Toggle led
					If rpmperiod > 200 And rpmperiod < 12000 Then
						rpmout = rpmperiod
					Endif
					If pulseperiod < 60000 Then
						pulseout = pulseperiod
					Endif
					Hserout #pulseout, ";", #rpmout, ";", #lambdaout, ";", Lf
			Else
					diagdisp = diagdisp + 1
			Endif

EndSelect

If getlambda = False And lambdawait > 900 Then
	lambdaperiod = 0
	lambdawait = 0
	getlambda = True
Else
	lambdaout = lambdaperiod
	lambdawait = lambdawait + 1
	PORTB.2 = False
Endif

Goto main_loop
End                                               

On Interrupt OC1A

If lambda = True Then getlambda = False

If getlambda = True Then
	PORTB.2 = True
	lambdaperiod = lambdaperiod + 1
Endif

If pulse = 0 Then  'sensor detected
		If pulsehighwait = False Then  'ignore noise
				pulsehighwait = True
				pulseperiod = pulselow + 1
				pulselow = 0
				pulsedisp = True
		Else
			pulselow = pulselow + 1  'count period
		Endif
Else
		pulselow = pulselow + 1  'count period
		pulsehighwait = False
Endif


If spark = 1 Then  'ignition detected !!!
		If rpmhighwait > 180 Then  'ignore noise
			rpmhighwait = 0
			rpmperiod = rpmlow + 1
			rpmlow = 0
		Else
			rpmlow = rpmlow + 1  'count period
			'rpmhighwait = 0
			rpmhighwait = rpmhighwait + 1
		Endif
Else
		rpmlow = rpmlow + 1  'count period
		rpmhighwait = rpmhighwait + 1
Endif



Resume                                            

