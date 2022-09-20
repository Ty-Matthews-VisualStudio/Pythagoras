How to compile Pythagoras
=======================

Pythagoras currently requires Visual Studio to compile.  A freely available community edition can be found on Microsoft's website (https://visualstudio.microsoft.com/vs/community/) At the time of writing (February 2020), Pythagoras was compiled using Visual Studio 2019 version.  Newer versions of VS should support opening older formats in case the 2019 version is no longer available.

Look for Pythagoras.sln in the root folder, this is the solution file that contains each of the individual projects.  There are five separate VC++ projects inside this solution:

1. Pythagoras: The main VC++/MFC application that the user interacts with.
2. PythagorasEngine: A Windows executable without any GUI that is spawned on demand by Pythagoras.  This engine is ultimately what calls Python scripts when the user executes them from within Pythagoras.
3. PythagorasInstaller: A dialog-based application that installs Pythagoras, Python, the Visual C++ redistributable, and some common Python libraries that the scripts use.
4. TestInterprocess: A debugging application for testing the Boost interprocess code.
5. zlibstat: A library for zipping and unzipping files.

Further details on each project are given below.

Dependencies
---------------------
The Pythagoras projects have two external library dependencies: Python and Boost.  The MFC and Windows API libraries come by default during the install of Visual Studio.  

*Boost*
The Boost library can be obtained from https://www.boost.org/.  It is predominantly a headers-only library, although there are precompiled binaries that are necessary for certain functions.  Both Pythagoras and PythagorasEngine depend on the Boost library, while the other three do not.  The Boost.org website has downloads for the headers-only files, while pre-compiled binaries for Windows can be located at https://sourceforge.net/projects/boost/files/boost-binaries/.  Both Pythagoras and PythagorasEngine require the precompiled binaries to link properly.  These libraries must be located in the <Boost>/stage/lib folder for the linker to find them (where <Boost> is the root of the specific version of Boost you have downloaded, e.g. Documents\Code\boost_1_72_0)

At the time of writing (February 2020), the Boost library version was 1.72.0.  Precompiled versions for Boost are available for different versions of Boost itself as well as different versions of the Visual C++ build tools (e.g. MSVC v142).  Make sure you download both x64 and x86 versions for the correct MSVC toolset you are using.  Alternatively, you can also build the precompiled static libraries using the Boost bootstrap tool, which will build everything from source code using the build environment you currently have, which ensures that your static libraries will be compatible.  More information for how to build manually can be found at the following URL:

https://www.boost.org/doc/libs/1_72_0/more/getting_started/windows.html

The command I used to build both the 32 and 64 bit versions of the Boost libraries is given below:

>b2.exe --toolset=msvc-14.2 --address-model=64 --architecture=x86 --runtime-link=static,shared --link=static threading=multi --build-dir=build --stage-dir=stage -j8

This command will generate the static lib files and place them into the boost/stage directory.

Once you have the Boost headers and static libraries downloaded (or built), you'll need to specify where they are located so the project can find them.  To do this, open the PropertyManager pane by going to View...PropteryManager within Visual Studio.  Open the Pythagoras twisty, then open the Release|x64 twisty.  Double-click the "PropertySheet" item to edit it.  Click on the "User Macros" entry from the PropertySheet to edit all of the Preprocessor definitions that have local path entries.  There is only one PropertySheet for all projects, and it is not platform-specifc, so you'll only have to edit one of them.  Change the path as shown for each of the entries to reflect your local directory structure.  These Preprocessor macros are then referenced within each project that uses them with $(Variable) type syntax.

*Python*
The Pythagoras project has build settings for both x86 and x64 targets.  The Python libraries for these two platforms are different, and they are not compatible.  Meaning, if you compile with the 32-bit version of the Python library, it will not work if the 64-bit version of Python is installed on the client (and vice-versa).  To remedy this, and to be able to support x86 clients, I have both x64 and x86 Python installations on the compile computer and I link to the appropriate one using the PropertySheet path information as described in the *Boost* section.  This is why there is a Pythonx64Dir and a Pythonx86Dir (the Boost library handles both platforms using headers that link to the correct versions so separate Prep).

If you plan to install both 32-bit and 64-bit versions of Python of the same computer, it is important that you not allow the installer to set the path to point to the version that you do not wish to use as the primary Python install.  For my desktop, this meant installing Python 64 first, letting it set the path correctly (and for all users), then installing Python 32-bit and deselecting the "set path" and "install for all users" options.  In this way, I kept Python 64 as the default version for everyday use but still had a 32-bit version for compiling purposes.

Python versioning goes by X.Y.Z, where X is the major version number, Y is the minor version number, and Z is the micro version number.  Python versions in the same micro level (e.g. 3.7.4 and 3.7.6) **should** be compatible with each other and not require recompiling with new libraries.  This is because both 3.7.4 and 3.7.6 would have a library named Python37.lib (the micro version is not included).  However, when the minor version changes it's probably going to require a recompile, since a new Python library will likely get issued when the minor version changes.  It is not clear whether a program compiled against 3.Y would work if the client has 3.Y+1 installed.  