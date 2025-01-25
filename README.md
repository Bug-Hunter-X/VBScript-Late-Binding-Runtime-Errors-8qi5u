# VBScript Late Binding Runtime Errors

This repository demonstrates a common issue in VBScript: runtime errors caused by late binding when accessing unsupported methods or properties of an object.

## The Problem
Late binding, while offering flexibility, can lead to unexpected runtime errors if an object doesn't support the method or property being called.  The `bug.vbs` file shows an example where attempting to call a non-existent method results in an error.

## The Solution
The solution involves using early binding or adding robust error handling.  `bugSolution.vbs` provides an example that utilizes error handling to gracefully manage potential errors instead of crashing the script.

## How to Run
1. Save the `bug.vbs` and `bugSolution.vbs` files.
2. Double-click the files to run them using a VBScript interpreter.