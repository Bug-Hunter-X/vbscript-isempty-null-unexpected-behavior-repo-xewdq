# VBScript IsEmpty(Null) Unexpected Behavior

This repository demonstrates an unexpected behavior in VBScript when using the `IsEmpty()` function with a `Null` value.  The `IsEmpty()` function is designed to check if a variable is uninitialized, but it does not behave consistently when encountering `Null` values.

The `bug.vbs` file illustrates the problem.  The `bugSolution.vbs` file proposes a solution.

## Bug Description
In VBScript,  `IsEmpty(Null)` returns `False`,  which can lead to unexpected results if you rely on `IsEmpty` to handle null or missing values.  This differs from how `IsEmpty` treats an uninitialized variable (which returns `True`). 

## Solution
The solution is to explicitly check for `Null` using `IsNothing` function before checking for `IsEmpty`. This ensures proper handling of both uninitialized variables and `Null` values.

## How to Reproduce
1. Open a VBScript editor or IDE.
2. Copy and paste the code from `bug.vbs`
3. Run the script. Observe the unexpected output (0).
4.  Copy and paste the code from `bugSolution.vbs`. 
5. Run the script. Observe the corrected output.