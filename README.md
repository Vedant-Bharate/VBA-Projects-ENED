# VBA Projects ENED
This repository contains Visual Basic for Applications (VBA) projects completed for my Engineering Education (ENED) coursework. Explore various assignments focusing on VBA syntax, data analysis, automation, and advanced techniques.

## About

In this repository, you'll find various VBA projects focusing on automating tasks, analyzing data, and enhancing productivity in the context of engineering education. From simple macros to complex scripts, these projects showcase the application of VBA within the realm of engineering education.

## Table of Contents

- [Activities / Homework](#activities-or-homework)
- [Challenges / CFU](#challenges-or-cfu)
- [Contributing](#contributing)

## Activities or Homework

This section includes the homework assignments given for my ENED coursework. These assignments serve as activities to learn and improve VBA skills and may require some time to complete.
See [Activities](Activities) folder for more details

## Challenges or CFU

In this section, you'll find multiple challenges that need to be completed within a time limit (15-20 minutes). These challenges are designed to test your understanding of VBA concepts and techniques in a fast-paced environment.

## Contributing

If you would like to contribute to this repository, feel free to submit pull requests with improvements, additional projects, or bug fixes. Your contributions are highly appreciated!

## License

This repository is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

Happy coding!

---

# VBA Working
VBA is a ackend software for microsoft applications.
The VBA Editor offers several windows designed for developing and debugging macros, each serving a distinct purpose:

1. **Project Window**:

   **Purpose:** This window presents all components within your VBA project, including modules, forms, class modules, and references to external libraries. It facilitates management of these components, allowing addition of new ones and easy navigation between them.

   **Importance:** The Project window provides a structural overview of your entire project, simplifying organization and access to different code segments.

2. **Property Inspector**:

   **Purpose:** The Property Inspector displays properties and their corresponding values for the currently selected object, be it a form, control, or module. It permits modification of these properties to customize appearance and behavior of objects.

   **Importance:** This window enables fine-tuning of functionality and aesthetics of forms and controls without direct code manipulation.

3. **Code Editor**:

   **Purpose:** This serves as the workspace for writing and editing VBA code. It offers features like syntax highlighting, auto-completion, and debugging tools to aid in code creation and troubleshooting.

   **Importance:** The Code Editor is the nucleus of the VBA development environment, where instructions are crafted to automate tasks and extend functionalities within Office applications.

4. **Immediate Window**:

   **Purpose:** The Immediate Window allows direct execution of VBA commands and immediate viewing of their outcomes. It's beneficial for testing code snippets, evaluating expressions, and debugging macros.

   **Importance:** Offering a swift and interactive means to test and debug code, the Immediate window obviates the need to run the entire macro or set breakpoints.

5. **Locals Window**:

   **Purpose:** This window exhibits values of all variables and objects within the current scope during code execution. It aids in comprehending data changes throughout the macro and identifying potential issues.

   **Importance:** As a valuable debugging tool, the Locals window permits inspection of variable and object states at specific code execution points.

* All code is written as subroutines in the Macro. any code you want to run in the macro should be enclosed within `sub` and `End Sub`.

## Variables Declarations
` Dim Variable As VariableType`
`integer` `long` `double` `string` `boolean`

* `CDbl()` - convert to type double

* `CInt()` - convert to type integer

* `CLng()` - convert to type long

* `CStr()` - convert to type string

* `CBool()` - convert to type boolean

## Performing Calculations:

Precedence
`^` Power
`+ve or -ve` Positive/Negative
`* or /` Multiplication/Division
`%`  Modulus
`+ or -` Addition/Subtraction

## Complex Mathematical Operations.
Trigonometric: `sin()`, `cos()`, `tan()`
Logarithmic: `log()`
Exponential: `exp()`
Absolute Value: `abs ()`
Square Root: `sqr()`
Rounding: `round()`

## Accessing the value of the worksheet
`var = ActiveSheet.Cells(Rows,Columns).Value`
* ActiveSheet tells the editor to access the value of the current open worksheet.
The **index** for the Rows and Columns starts with 1.

## Saving and Opening the Excel Files.
The excel should be saved in the excel Macro-Enabled Workbook. (.xlsm)

## Printing Values

* `Debug.Print x` would print the value of x in the **immediate window**.
* `MsgBox x` would print the value of x in the **Message Box ( Alert )** Pop up Alert.


## Using the Conditional Operators and statements.
* Value1 `=` Value2 (Equals to)
* Value1 `<>` Value2 (Not Equals to)
* Value1 `<` Value2 (Lessthan)
* Value1 `<=` Value2 (Less than Equal to)
* Value1 `>` Value2 (Greater than)
* Value1 `>=` Value2 (Greater than Equal to)

* Value1 `And` Value2
* Value1 `Or` Value2
* `Not`(Value1)

The operators used above work for numbers, single characters and strings. the ASCII values are used to campare these values and then finding the corresponding value. Two strings are said to be equal if they have the same characters in the same pattern. You can determine the order of strings by using the above operators.

` "Hear" < "Here" `
The result will be True as hear is precent to here in the alphabetical order.

## If else Ladder.

* Syntax:
```
If condition Then
 conditional code
ElseIf condition Then
 conditional code
Else 
 else code
End If
```

# Loops
* Conditional Based Reptition: **More Versatile** The identified actions are continued untill some event occurs that causes the algorithm to stop repeating.

* Counter Based Repetition: The identified actions are repeated for a specific number of times determined prior to starting the repeatition.


## For.... Next Loop

**First Method**
```
For var = start to End Step Change
   ....code to repeat....
Next
```
**Second method**
```
For Each Item In Array
   ....code to repeat....
Next
```


## Initialization of Array

```
Dim ArrayName(N) as VariableType
Dim 2DArray(R, C) as VariableType

ReDim ArrayName(N) as VariableType
ReDim Preserve ArrayName(N) as VariableType
```



## Functions in VBA

```Function FunctionName(Parameter) As Type
   function code
   FunctionName = value to return
   
   end Function
```

### **Options for writing Funtions**

No parameters:
`Function MyFunction() As Type`

• One parameter:
`Function MyFunction (var As Type) As Type`

• Multiple parameters:
`Function MyFunction(varl As Type, var2 As Type,..) As Type`

• Array parameters:
`Function MyFunction (ArrayVar() As Type) As Type`

### **Ranges of Cells**
• As was mentioned previously, you can use a function that you've written either within a macro or within an Excel formula

• The only change that needs to be made in writing the function to use it in an Excel formula is if you need to access a **range of cells**

• If you need to access a single cell or several single cells, you simply create a parameter for each value needed:

`Function MyFunction (Celll As Type, Cell2 As Type,...) As Type`

• If you need to access a range of cells, instead of specifying the parameter as an array, you specify the parameter as a range:

`Function MyFunction(CellValues As Range) As Type`
```
Function Avg2 (v As Range) As Double
Dim k As Integer
Dim Sum As Double
Dim Cell As Range
Dim Count As Integer
Sum = 0
Count = 0
For Each Cell In
Sum = Sum + Cell.Value
Count = Count + 1
Next
Avg2 = Sum / Count
End
```

### **Returning Values**

• Unlike functions in MATLAB and Python, a function in VBA can only return a single value or an array of values
• To return a single value from a function, we use the syntax we've been using:
`Function MyFunction (parameters) As Type`

• To return an array of values from a function, we need to specify the type of the function to be an array:
`Function MyFunction (parameters) As Type()`
