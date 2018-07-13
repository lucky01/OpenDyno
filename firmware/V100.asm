; Compiled with: AVR Simulator IDE v2.29
; Microcontroller model: ATtiny2313
; Clock frequency: 4.0MHz
;
.EQU    XHL = 0x01A
.EQU    ZHL = 0x01E
.EQU    R1716 = 0x010
.EQU    R1918 = 0x012
.EQU    R2120 = 0x014
.EQU    R19181716 = 0x010
;       The address of 'bttimeout' (word) (global) is 0x067
;       The address of 'pulsedisp' (bit) (global) is 0x019,0
;       The address of 'diagdisp' (word) (global) is 0x069
;       The address of 'rpm' (word) (global) is 0x071
;       The address of 'rpmperiod' (word) (global) is 0x00A
;       The address of 'rpmout' (word) (global) is 0x064
;       The address of 'rpmhighwait' (word) (global) is 0x01C
;       The address of 'rpmlow' (word) (global) is 0x00C
;       The address of 'pulseperiod' (word) (global) is 0x060
;       The address of 'pulseout' (word) (global) is 0x06B
;       The address of 'pulsehighwait' (bit) (global) is 0x019,1
;       The address of 'pulselow' (word) (global) is 0x017
;       The address of 'charinput' (byte) (global) is 0x066
;       The address of 'mode' (byte) (global) is 0x00E
;       The address of 'getlambda' (bit) (global) is 0x019,2
;       The address of 'lambdaperiod' (word) (global) is 0x062
;       The address of 'lambdawait' (word) (global) is 0x06D
;       The address of 'lambdaout' (word) (global) is 0x06F
;       The address of 'bt1' (bit) (global) is 0x036,4
;       The address of 'bt2' (bit) (global) is 0x036,5
;       The address of 'bt3' (bit) (global) is 0x036,6
;       The address of 'pulse' (bit) (global) is 0x030,2
;       The address of 'led' (bit) (global) is 0x032,4
;       The address of 'spark' (bit) (global) is 0x030,3
;       The address of 'lambda' (bit) (global) is 0x028,5
.ORG	0x000000
	RJMP L0002
.ORG	0x000004
	RJMP L0003
L0004:
	POP R31
	POP R30
	POP R27
	POP R26
	POP R22
	POP R21
	POP R20
	POP R19
	POP R18
	POP R17
	POP R16
	POP R15
	POP R0
	POP R16
	OUT SREG,R16
	POP R16
	RETI
L0002:
	CLR R15
	LDI R16,low RAMEND
	OUT SPL,R16
	RJMP L0005
; User code start
L0005:
; 1: Dim bttimeout As Word
; 2: bttimeout = 0
	STS 0x067,R15
	STS 0x068,R15
; 3: 
; 4: Dim pulsedisp As Bit
; 5: Dim diagdisp As Word
; 6: 
; 7: pulsedisp = False
	CBR 0x019,1
; 8: diagdisp = 0
	STS 0x069,R15
	STS 0x06A,R15
; 9: 
; 10: Dim rpm As Word
; 11: Dim rpmperiod As Word
; 12: Dim rpmout As Word
; 13: Dim rpmhighwait As Word
; 14: Dim rpmlow As Word
; 15: 
; 16: rpm = 0
	STS 0x071,R15
	STS 0x072,R15
; 17: rpmperiod = 0
	MOV 0x00A,R15
	MOV 0x00B,R15
; 18: rpmout = 0
	STS 0x064,R15
	STS 0x065,R15
; 19: rpmhighwait = 0
	LDI 0x01C,0x00
	LDI 0x01D,0x00
; 20: rpmlow = 0
	MOV 0x00C,R15
	MOV 0x00D,R15
; 21: 
; 22: Dim pulseperiod As Word
; 23: Dim pulseout As Word
; 24: Dim pulsehighwait As Bit
; 25: Dim pulselow As Word
; 26: 
; 27: pulseout = 0
	STS 0x06B,R15
	STS 0x06C,R15
; 28: pulseperiod = 0
	STS 0x060,R15
	STS 0x061,R15
; 29: pulsehighwait = 0
	CBR 0x019,2
; 30: pulselow = 0
	LDI 0x017,0x00
	LDI 0x018,0x00
; 31: 
; 32: Dim charinput As Byte
; 33: charinput = 0
	STS 0x066,R15
; 34: Dim mode As Byte
; 35: mode = 0
	MOV 0x00E,R15
; 36: 
; 37: ConfigPin PORTD.2 = Input  'pulse
	CBI DDRD,2
; 38: ConfigPin PORTD.3 = Input  'spark
	CBI DDRD,3
; 39: ConfigPin PORTD.4 = Output  'LED
	SBI DDRD,4
; 40: ConfigPin PORTB.0 = Input  'reference
	CBI DDRB,0
; 41: ConfigPin PORTB.1 = Input  'lambda input
	CBI DDRB,1
; 42: ConfigPin PORTB.2 = Output  'load Capacitor
	SBI DDRB,2
; 43: ConfigPin PORTB.4 = Input  'button 1
	CBI DDRB,4
; 44: ConfigPin PORTB.5 = Input  'button 2
	CBI DDRB,5
; 45: ConfigPin PORTB.6 = Input  'button 3
	CBI DDRB,6
; 46: 
; 47: PORTD.2 = 0
	CBI PORTD,2
; 48: PORTD.3 = 0
	CBI PORTD,3
; 49: PORTB.0 = 0
	CBI PORTB,0
; 50: PORTB.1 = 0
	CBI PORTB,1
; 51: PORTB.4 = 0
	CBI PORTB,4
; 52: PORTB.5 = 0
	CBI PORTB,5
; 53: PORTB.6 = 0
	CBI PORTB,6
; 54: 
; 55: Symbol bt1 = PINB.4
; 56: Symbol bt2 = PINB.5
; 57: Symbol bt3 = PINB.6
; 58: Symbol pulse = PIND.2
; 59: Symbol led = PORTD.4
; 60: Symbol spark = PIND.3
; 61: Symbol lambda = ACSR.ACO
; 62: 
; 63: ACSR.ACD = 0
	CBI ACSR,7
; 64: Disable ACI
	CBI ACSR,3
; 65: 
; 66: 'set timer parameters
; 67: TIMSK.OCIE1A = 1
	SET
	IN R16,TIMSK
	BLD R16,6
	OUT TIMSK,R16
; 68: TCCR1B.CS12 = 0
	CLT
	IN R16,TCCR1B
	BLD R16,2
	OUT TCCR1B,R16
; 69: TCCR1B.CS11 = 0
	CLT
	IN R16,TCCR1B
	BLD R16,1
	OUT TCCR1B,R16
; 70: TCCR1B.CS10 = 1
	SET
	IN R16,TCCR1B
	BLD R16,0
	OUT TCCR1B,R16
; 71: TCCR1B.CTC1 = 1
	SET
	IN R16,TCCR1B
	BLD R16,3
	OUT TCCR1B,R16
; 72: OCR1AH = $00
	OUT OCR1AH,R15
; 73: OCR1AL = $63
	LDI R16,0x63
	OUT OCR1AL,R16
; 74: Enable OC1A
	SET
	IN R16,TIMSK
	BLD R16,6
	OUT TIMSK,R16
; 75: Enable
	SEI
; 76: 
; 77: 
; 78: Dim getlambda As Bit
; 79: Dim lambdaperiod As Word
; 80: Dim lambdawait As Word
; 81: Dim lambdaout As Word
; 82: 
; 83: getlambda = False
	CBR 0x019,4
; 84: lambdaperiod = 0
	STS 0x062,R15
	STS 0x063,R15
; 85: lambdawait = 0
	STS 0x06D,R15
	STS 0x06E,R15
; 86: lambdaout = 0
	STS 0x06F,R15
	STS 0x070,R15
; 87: 
; 88: Hseropen 57600
; exact baud rate achieved = 55555.55; bit period = 18µs; baud rate error = 3.54%
	LDI R16,0x08
	OUT UBRRL,R16
	LDI R16,0x00
	OUT UBRRH,R16
	LDI R16,0x42
	OUT UCSRA,R16
	LDI R16,0x06
	OUT UCSRC,R16
	LDI R16,0x18
	OUT UCSRB,R16
; 89: led = True
	SBI PORTD,4
; 90: 
; 91: main_loop:
L0001:
; 92: 
; 93: Hserget charinput
	MOV R31,R15
	SBIC UCSRA,RXC
	IN R31,UDR
	STS 0x066,R31
; 94: 
; 95: Select Case charinput
; 96: Case "0"
	LDS R16,0x066
	LDI R17,0x30
	CPSE R16,R17
	RJMP L0006
; 97: mode = 0
	MOV 0x00E,R15
; 98: Hserout "STOP", Lf
	LDI R31,0x53
	RCALL HS01
	LDI R31,0x54
	RCALL HS01
	LDI R31,0x4F
	RCALL HS01
	LDI R31,0x50
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 99: led = True
	SBI PORTD,4
; 100: Case "1"
	RJMP L0007
L0006:
	LDS R16,0x066
	LDI R17,0x31
	CPSE R16,R17
	RJMP L0008
; 101: mode = 1
	LDI R16,0x01
	MOV 0x00E,R16
; 102: Hserout "DATA", Lf
	LDI R31,0x44
	RCALL HS01
	LDI R31,0x41
	RCALL HS01
	LDI R31,0x54
	RCALL HS01
	LDI R31,0x41
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 103: led = False
	CBI PORTD,4
; 104: Case "2"
	RJMP L0009
L0008:
	LDS R16,0x066
	LDI R17,0x32
	CPSE R16,R17
	RJMP L0010
; 105: mode = 2
	LDI R16,0x02
	MOV 0x00E,R16
; 106: Hserout "DIAG", Lf
	LDI R31,0x44
	RCALL HS01
	LDI R31,0x49
	RCALL HS01
	LDI R31,0x41
	RCALL HS01
	LDI R31,0x47
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 107: led = False
	CBI PORTD,4
; 108: EndSelect
L0010:
L0009:
L0007:
; 109: 
; 110: If bttimeout = 0 Then
	LDS R16,0x067
	LDS R17,0x068
	LDI R18,0x00
	LDI R19,0x00
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRS R20,SREG_Z
	RJMP L0011
; 111: If bt1 = 0 Then
	SBIC PINB,4
	RJMP L0012
; 112: Hserout "BT1", Lf
	LDI R31,0x42
	RCALL HS01
	LDI R31,0x54
	RCALL HS01
	LDI R31,0x31
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 113: Endif
L0012:
; 114: If bt2 = 0 Then
	SBIC PINB,5
	RJMP L0013
; 115: Hserout "BT2", Lf
	LDI R31,0x42
	RCALL HS01
	LDI R31,0x54
	RCALL HS01
	LDI R31,0x32
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 116: Endif
L0013:
; 117: If bt3 = 0 Then
	SBIC PINB,6
	RJMP L0014
; 118: Hserout "BT3", Lf
	LDI R31,0x42
	RCALL HS01
	LDI R31,0x54
	RCALL HS01
	LDI R31,0x33
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 119: Endif
L0014:
; 120: bttimeout = 5000
	LDI R16,0x88
	STS 0x067,R16
	LDI R16,0x13
	STS 0x068,R16
; 121: Else
	RJMP L0015
L0011:
; 122: bttimeout = bttimeout - 1
	LDS R16,0x067
	LDS R17,0x068
	LDI R18,0x01
	LDI R19,0x00
	SUB R16,R18
	SBC R17,R19
	STS 0x067,R16
	STS 0x068,R17
; 123: Endif
L0015:
; 124: 
; 125: 
; 126: If pulselow >= 60000 Then
	MOV R16,0x017
	MOV R17,0x018
	LDI R18,0x60
	LDI R19,0xEA
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRC R20,SREG_C
	RJMP L0016
; 127: pulselow = 60000
	LDI 0x017,0x60
	LDI 0x018,0xEA
; 128: pulseperiod = 0
	STS 0x060,R15
	STS 0x061,R15
; 129: pulseout = 0
	STS 0x06B,R15
	STS 0x06C,R15
; 130: Endif
L0016:
; 131: 
; 132: If rpmlow >= 60000 Then
	MOV R16,0x00C
	MOV R17,0x00D
	LDI R18,0x60
	LDI R19,0xEA
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRC R20,SREG_C
	RJMP L0017
; 133: rpmout = 0
	STS 0x064,R15
	STS 0x065,R15
; 134: rpmperiod = 0
	MOV 0x00A,R15
	MOV 0x00B,R15
; 135: rpmlow = 60000
	LDI R16,0x60
	MOV 0x00C,R16
	LDI R16,0xEA
	MOV 0x00D,R16
; 136: Endif
L0017:
; 137: 
; 138: If lambdaperiod >= 50000 Then
	LDS R16,0x062
	LDS R17,0x063
	LDI R18,0x50
	LDI R19,0xC3
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRC R20,SREG_C
	RJMP L0018
; 139: lambdaperiod = 50000
	LDI R16,0x50
	STS 0x062,R16
	LDI R16,0xC3
	STS 0x063,R16
; 140: lambdaout = 0
	STS 0x06F,R15
	STS 0x070,R15
; 141: Endif
L0018:
; 142: 
; 143: Select Case mode
; 144: Case 1
	MOV R16,0x00E
	LDI R17,0x01
	CPSE R16,R17
	RJMP L0019
; 145: If pulsedisp = True And pulseperiod < 60000 Then
	SBRS 0x019,0
	RJMP L0020
	LDS R16,0x060
	LDS R17,0x061
	LDI R18,0x60
	LDI R19,0xEA
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0020
; 146: Toggle led
	SBI PIND,4
	SBI DDRD,4
; 147: If rpmperiod > 240 And rpmperiod < 12000 Then
	MOV R16,0x00A
	MOV R17,0x00B
	LDI R18,0xF0
	LDI R19,0x00
	CP R18,R16
	CPC R19,R17
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0021
	MOV R16,0x00A
	MOV R17,0x00B
	LDI R18,0xE0
	LDI R19,0x2E
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0021
; 148: rpmout = rpmperiod
	STS 0x064,0x00A
	STS 0x065,0x00B
; 149: Endif
L0021:
; 150: Hserout #pulseperiod, ";", #rpmout, ";", #lambdaout, ";", Lf
	LDS 0x007,0x060
	LDS 0x008,0x061
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDS 0x007,0x064
	LDS 0x008,0x065
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDS 0x007,0x06F
	LDS 0x008,0x070
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 151: pulsedisp = False
	CBR 0x019,1
; 152: Endif
L0020:
; 153: Case 2
	RJMP L0022
L0019:
	MOV R16,0x00E
	LDI R17,0x02
	CPSE R16,R17
	RJMP L0023
; 154: If diagdisp >= 1000 Then
	LDS R16,0x069
	LDS R17,0x06A
	LDI R18,0xE8
	LDI R19,0x03
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRC R20,SREG_C
	RJMP L0024
; 155: diagdisp = 0
	STS 0x069,R15
	STS 0x06A,R15
; 156: Toggle led
	SBI PIND,4
	SBI DDRD,4
; 157: If rpmperiod > 200 And rpmperiod < 12000 Then
	MOV R16,0x00A
	MOV R17,0x00B
	LDI R18,0xC8
	LDI R19,0x00
	CP R18,R16
	CPC R19,R17
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0025
	MOV R16,0x00A
	MOV R17,0x00B
	LDI R18,0xE0
	LDI R19,0x2E
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0025
; 158: rpmout = rpmperiod
	STS 0x064,0x00A
	STS 0x065,0x00B
; 159: Endif
L0025:
; 160: If pulseperiod < 60000 Then
	LDS R16,0x060
	LDS R17,0x061
	LDI R18,0x60
	LDI R19,0xEA
	CP R16,R18
	CPC R17,R19
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0026
; 161: pulseout = pulseperiod
	LDS R16,0x060
	STS 0x06B,R16
	LDS R16,0x061
	STS 0x06C,R16
; 162: Endif
L0026:
; 163: Hserout #pulseout, ";", #rpmout, ";", #lambdaout, ";", Lf
	LDS 0x007,0x06B
	LDS 0x008,0x06C
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDS 0x007,0x064
	LDS 0x008,0x065
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDS 0x007,0x06F
	LDS 0x008,0x070
	RCALL _append_lab_0001
	LDI XL,0x01
	LDI XH,0x00
	RCALL HS21
	LDI R31,0x3B
	RCALL HS01
	LDI R31,0x0A
	RCALL HS01
; 164: Else
	RJMP L0027
L0024:
; 165: diagdisp = diagdisp + 1
	LDS R16,0x069
	LDS R17,0x06A
	INC R16
	BRNE PC+2
	INC R17
	STS 0x069,R16
	STS 0x06A,R17
; 166: Endif
L0027:
; 167: 
; 168: EndSelect
L0023:
L0022:
; 169: 
; 170: If getlambda = False And lambdawait > 900 Then
	SBRC 0x019,2
	RJMP L0028
	LDS R16,0x06D
	LDS R17,0x06E
	LDI R18,0x84
	LDI R19,0x03
	CP R18,R16
	CPC R19,R17
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0028
; 171: lambdaperiod = 0
	STS 0x062,R15
	STS 0x063,R15
; 172: lambdawait = 0
	STS 0x06D,R15
	STS 0x06E,R15
; 173: getlambda = True
	SBR 0x019,4
; 174: Else
	RJMP L0029
L0028:
; 175: lambdaout = lambdaperiod
	LDS R16,0x062
	STS 0x06F,R16
	LDS R16,0x063
	STS 0x070,R16
; 176: lambdawait = lambdawait + 1
	LDS R16,0x06D
	LDS R17,0x06E
	INC R16
	BRNE PC+2
	INC R17
	STS 0x06D,R16
	STS 0x06E,R17
; 177: PORTB.2 = False
	CBI PORTB,2
; 178: Endif
L0029:
; 179: 
; 180: Goto main_loop
	RJMP L0001
; 181: End
L0030	RJMP L0030
; 182: 
; 183: On Interrupt OC1A
L0003:
	PUSH R16
	IN R16,SREG
	PUSH R16
	PUSH R0
	PUSH R15
	PUSH R16
	PUSH R17
	PUSH R18
	PUSH R19
	PUSH R20
	PUSH R21
	PUSH R22
	PUSH R26
	PUSH R27
	PUSH R30
	PUSH R31
; 184: 
; 185: If lambda = True Then getlambda = False
	SBIS ACSR,5
	RJMP L0031
	CBR 0x019,4
L0031:
; 186: 
; 187: If getlambda = True Then
	SBRS 0x019,2
	RJMP L0032
; 188: PORTB.2 = True
	SBI PORTB,2
; 189: lambdaperiod = lambdaperiod + 1
	LDS R16,0x062
	LDS R17,0x063
	INC R16
	BRNE PC+2
	INC R17
	STS 0x062,R16
	STS 0x063,R17
; 190: Endif
L0032:
; 191: 
; 192: If pulse = 0 Then  'sensor detected
	SBIC PIND,2
	RJMP L0033
; 193: If pulsehighwait = False Then  'ignore noise
	SBRC 0x019,1
	RJMP L0034
; 194: pulsehighwait = True
	SBR 0x019,2
; 195: pulseperiod = pulselow + 1
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x017
	ADC R17,0x018
	STS 0x060,R16
	STS 0x061,R17
; 196: pulselow = 0
	LDI 0x017,0x00
	LDI 0x018,0x00
; 197: pulsedisp = True
	SBR 0x019,1
; 198: Else
	RJMP L0035
L0034:
; 199: pulselow = pulselow + 1  'count period
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x017
	ADC R17,0x018
	MOV 0x017,R16
	MOV 0x018,R17
; 200: Endif
L0035:
; 201: Else
	RJMP L0036
L0033:
; 202: pulselow = pulselow + 1  'count period
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x017
	ADC R17,0x018
	MOV 0x017,R16
	MOV 0x018,R17
; 203: pulsehighwait = False
	CBR 0x019,2
; 204: Endif
L0036:
; 205: 
; 206: 
; 207: If spark = 1 Then  'ignition detected !!!
	SBIS PIND,3
	RJMP L0037
; 208: If rpmhighwait > 180 Then  'ignore noise
	MOV R16,0x01C
	MOV R17,0x01D
	LDI R18,0xB4
	LDI R19,0x00
	CP R18,R16
	CPC R19,R17
	IN R20,SREG
	SBRS R20,SREG_C
	RJMP L0038
; 209: rpmhighwait = 0
	LDI 0x01C,0x00
	LDI 0x01D,0x00
; 210: rpmperiod = rpmlow + 1
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x00C
	ADC R17,0x00D
	MOV 0x00A,R16
	MOV 0x00B,R17
; 211: rpmlow = 0
	MOV 0x00C,R15
	MOV 0x00D,R15
; 212: Else
	RJMP L0039
L0038:
; 213: rpmlow = rpmlow + 1  'count period
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x00C
	ADC R17,0x00D
	MOV 0x00C,R16
	MOV 0x00D,R17
; 214: 'rpmhighwait = 0
; 215: rpmhighwait = rpmhighwait + 1
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x01C
	ADC R17,0x01D
	MOV 0x01C,R16
	MOV 0x01D,R17
; 216: Endif
L0039:
; 217: Else
	RJMP L0040
L0037:
; 218: rpmlow = rpmlow + 1  'count period
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x00C
	ADC R17,0x00D
	MOV 0x00C,R16
	MOV 0x00D,R17
; 219: rpmhighwait = rpmhighwait + 1
	LDI R16,0x01
	LDI R17,0x00
	ADD R16,0x01C
	ADC R17,0x01D
	MOV 0x01C,R16
	MOV 0x01D,R17
; 220: Endif
L0040:
; 221: 
; 222: 
; 223: 
; 224: Resume
	RJMP L0004
; 225: 
; End of user code
L0041	RJMP L0041
; APPEND CODE BEGIN: _routine_ascii_word_
_append_lab_0001:
	LDI zl,0x01
	MOV zh,r15
	clt
	MOV R20,0x007
	MOV R21,0x008
	LDI R18,0x10
	LDI R19,0x27
	RCALL _append_lab_0002
	LDI R18,0xE8
	LDI R19,0x03
	RCALL _append_lab_0002
	LDI R18,0x64
	LDI R19,0x00
	RCALL _append_lab_0002
	LDI R18,0x0A
	LDI R19,0x00
	RCALL _append_lab_0002
	MOV R0,R20
	RCALL _append_lab_0003
	ST z,r15
	RET
_append_lab_0002:
	MOV R16,R20
	MOV R17,R21
	RCALL D001
	MOV r0,r16
	CPSE r0,r15
	set
	BRTC _append_lab_0004
_append_lab_0003:
	LDI r16,0x30
	ADD r0,r16
	ST z+,r0
_append_lab_0004:
	RET
; APPEND CODE END.
;
;
; Word Division Routine
D001:	CLR R20
	SUB R21,R21
	LDI R22,0x11
D002:	ROL R16
	ROL R17
	DEC R22
	BRNE D003
	RET
D003:	ROL R20
	ROL R21
	SUB R20,R18
	SBC R21,R19
	BRCC D004
	ADD R20,R18
	ADC R21,R19
	CLC
	RJMP D002
D004:	SEC
	RJMP D002
; Hardware Serial Communication Routines
HS01:	SBIS UCSRA,UDRE
	RJMP HS01
	OUT UDR,R31
	RET
HS10:	SBIS UCSRA,RXC
	RJMP HS10
	IN R31,UDR
	RET
; Hserout Decimal Conversion Routine
HS22:
	MOVW XH:XL,R1:R0
HS21:	LD R31,X+
	MOVW R1:R0,XH:XL
	ANDI R31,0xFF
	BRNE PC+2
	RET
	RCALL HS01
	RJMP HS22
;
;
; End of listing
.END
