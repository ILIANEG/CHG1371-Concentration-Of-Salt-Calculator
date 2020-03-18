Option Explicit

''Variables that are gonna be used throughout the module
''initial will become the array of initial values used at every step as well as return array
Dim initial As Variant
''formulas will become an array of coefficients
Dim formulas As Variant
''k will become an array of k-values
Dim k As Variant
''h is a step that used to increment x as well as used in calculation of k-values
Dim h As Double
''variables is an array of Csn values incremented by h * SUM(BjKj) from RK4 formula
Dim variables As Variant

''Function which receives initial condition of system (i), matrix of coefficients (f)
''Values of initial and final time (initialFinalT) and step size used to calculate all intermediates
Public Function FINDCONCENTRATION(i As Range, f As Range, initialFinalT As Range, step As Double) As Variant
	''integers used in For loops
	Dim a As Integer
	Dim b As Integer

	''Assignment of module variable h, representing increment step
	h = step
	''Turning initial into 1-D array
	ReDim initial(1 To i.Columns.Count)
	''Turning formulas into 2-D array
	ReDim formulas(1 To f.Rows.Count, 1 To f.Columns.Count)

	''Populating initial values array
	For a = 1 To i.Columns.Count
		initial(a) = i(1, a).Value
	Next a

	''Populating formula cofficients array
	For a = 1 To f.Rows.Count
		For b = 1 To f.Columns.Count
			formulas(a, b) = f(a, b).Value
		Next b
	Next a

	''Since change of concentration/time in any tank is not explivitely dependant on time
	''In other words variable "t" is not element of any equation we can imagine that we start from 0
	''every time, and going to some final x value with certain increment, in this case we going to 1
	''by increment of h
	Dim seekX As Double, currX As Double
	seekX = initialFinalT(2, 1).Value - initialFinalT(1, 1).Value
	currX = 0

	''Engine of the function, loop receives array of k values, turns it into Yn+1
	''and uses new values as initial
	Do Until currX > seekX
		Call kCalc
		For a = 1 To length(initial, 1)
			initial(a) = calcY(a)
		Next a
		currX = currX + h
	Loop
	''returns answer
	FINDCONCENTRATION = initial
End Function

Private Function kCalc()

	''i, j, k used in nested loop
	Dim i As Integer
	Dim j As Integer
	Dim z As Integer

	''turns k into 2-D array
	ReDim k(1 To 4, 1 To length(formulas, 1))
	''turns variables into 1-D array
	ReDim variables(1 To length(formulas, 1))

	''Nested loop which iterates 4 times creating all k for every function at every level
	For i = 1 To 4
		''Populates variables with incremented values ready to be fed into the function
		For z = 1 To length(formulas, 1)
			If i = 1 Then
				variables(z) = initial(z)
			ElseIf i = 2 Or i = 3 Then
				variables(z) = initial(z) + (1 / 2) * h * k(i - 1, z)
			ElseIf i = 4 Then
				variables(z) = initial(z) + h * k(i - 1, z)
			End If
		Next z
		''Feeds variables into the function
		For j = 1 To length(formulas, 1)
			k(i, j) = calcF(variables, j)
		Next j
	Next i
End Function

''This function uses assumption of constant multipliers to multiply every variable by a coefficient
''inputed by user
Private Function calcF(variables As Variant, formulaIndex As Integer) As Double
	Dim i As Integer
	Dim res As Double
	res = 0
	For i = 1 To length(variables, 1)
		res = res + formulas(formulaIndex, i) * variables(i)
	Next i
	calcF = res
End Function

''Calculates Yn+1 based on collection of K-values using RK4
Private Function calcY() As Variant
	ReDim y(length(initial, 1)) As Variant
	Dim b As Integer
	For b = 1 To length(initial, 1)
		y(b) = initial(b) + h * (1 / 3) * ((1 / 2) * k(1, b) + k(2, b) + k(3, b) + (1 / 2) * k(4, b))
	Next b
	calcY = y
End Function

''Supporting function that calucates length of specified dimention of array, very usefull:)
Private Function length(i As Variant, d As Integer) As Integer
	length = UBound(i, d) - LBound(i, d) + 1
End Function
