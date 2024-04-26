# ExcelWrapper
Late-bound .NET object for interacting with Excel.

Supports common tasks for manipulating Excel data without requiring a reference to interop COM components.

## Features
- Supports both .NET Core and Framework.
- Straight-forward get and set methods for values and formulas - specify by location (e.g. A1, B23, BA21, etc.) or cell row/column coordinates.
- Methods for extracting values from cells using generics.
- Full support for creating, opening, and saving workbooks.
- Full support for adding, deleting, and moving worksheets.
- Support for common range operations such as setting borders, fonts, colors, and text formatting.
- Support for calling the native Excel Find function.
- Convenience methods for common operations (column number to letter and vice versa, formula reference builder (R1C1 format), etc.).
- Late-bound, so Excel interop libraries are not required to be referenced.
- Direct access to the underlying Excel COM objects to invoke any non-implemented functions.
- Implements `IDisposable` to prevent 'ghost' Excel instances.
- Full intellisense comments included.

_Note:_ While explicit references to Excel libraries are not required, this component does require that Excel is installed on any system which utilizes it. This library is not a replacement for Excel, it is a late-bound interface to the Excel COM objects.

## Usage Sample
See here: [Test Application/Program.cs](Test%20Application/Program.cs)
