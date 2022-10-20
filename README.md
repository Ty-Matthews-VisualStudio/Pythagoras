# Pythagoras
Pythagoras is an ETL application written in Visual C++ with embedded Python functionality.  The primary objective of Pythagoras is to provide the scientific user a method for quickly and easily processing and visualizing data gathered from a variety of scientific instruments that are used on a routine basis.  The application consists of two main parts: a Visual C++ GUI which provides the means to find and select files for processing, as well as choosing which type of data processing to run on those files.  The second part consists of a Python engine which executes Python scripts which do the heavy lifting of actually opening data files, processing the data, and generating the output.  The user has the option to create their own Python scripts for working with data, or to modify any existing script to suit their needs.

A typical use case is the following.  George has acquired some tensile data for some samples of a new polymer film he is developing.  He would like to be able to process the raw data in a routine way such that the output is generated in a controlled and repeatable way.  He opens Pythagoras, finds the files he recently generated, and then executes the desired function against those files.  The output in this case may be a calculated value of Young's Modulus and yield stress, along with a plot of the stress-strain curve.  The first thing George opens is Pythagoras:

![Pythagoras main screen](/Documentation/Pythagoras.PNG)

There are four main areas of the application:

1. The script browser which displays .py files in the local and common directory structures as well as a list of all executable functions inside the selected script
2. A list of files to be processed using the selected function
3. Window for selecting files, either through the Explorer tab or the RegEx search tab
4. Output window
