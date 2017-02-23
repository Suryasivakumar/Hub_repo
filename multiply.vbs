Option Explicit


Dim num1, num2, answer

'Set Variables
	num1 = InputBox ("Type the first number.")
	num2 = InputBox ("Type the second number.")
	
'Multiply the 2 numbers

	answer = num1 * num2
	
	MsgBox(num1 & " x " & num2 & " = " & answer)