Excel-Dna Documentation
=======================

Guide
^^^^^

Excel-DNA is an independent project to integrate .NET into Excel. 
With Excel-DNA you can make native (.xll) add-ins for Excel using C#, Visual Basic.NET or F#, providing high-performance user-defined functions (UDFs), 
custom ribbon interfaces and more. Your entire add-in can be packed into a single .xll file requiring no installation or registration.


More Details
============

Excel-DNA is developed using .NET, and users have to install the freely available .NET Framework runtime. 
The integration is by an Excel Add-In (.xll) that exposes .NET code to Excel. 
The user code can be in text-based (.dna) script files (C#, Visual Basic or F#), or compiled .NET libraries (.dll). 
Excel-DNA supports both the .NET runtime version 2.0 (which is used by .NET versions 2.0, 3.0 and 3.5) and version 4 (used by all newer versions). 
Add-ins can target either generation of the runtime, and concurrent loading of both runtime versions into an Excel instance is supported.

Excel versions ’97 through 2019 can be targeted with a single add-in. 
Advanced Excel features are supported, including multi-threaded recalculation (Excel 2007+), registration-free RTD servers (Excel 2002+) and customized Ribbon interfaces (Excel 2007+). 
There is support for integrated Custom Task Panes (Excel 2007+), offloading UDF computations to a Windows HPC cluster (Excel 2010+), and for the 64-bit versions of (Excel 2010+).

Most managed UDF assemblies developed for Excel Services can be exposed to the Excel client with no modification. 
(Please contact me if you are interested in this feature.)

Since Excel-DNA uses the Excel C API, porting C/C++ add-in code based on the Excel XLL SDK is very easy. (No more XLOPERs!)

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

Latest Releases
===============

The current version is Excel-DNA 1.0.0, released on 27 April 2019. This release included various bug-fixes and improvements to the NuGet package build integration. The Excel-DNA 1.0.x series will be the final versions with support for older .NET (< 4.0) and Excel (< 2007) releases.

.. toctree::
   :maxdepth: 2
   :caption: Contents:

   installation
   license
   help

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
