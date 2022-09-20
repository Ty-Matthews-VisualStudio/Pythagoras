// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>

#ifdef PYTHON32BIT
#ifdef _DEBUG
#undef _DEBUG
#include "c:\\Program Files (x86)\\Python36-32bit\\include\\Python.h"
#define _DEBUG
#else
#include "c:\\Program Files (x86)\\Python36-32bit\\include\\Python.h"
#endif
#else
#ifdef _DEBUG
#undef _DEBUG
#include "C:\\Users\\a4lg8zz\\Python\\include\\Python.h"
#define _DEBUG
#else
#include "C:\\Users\\a4lg8zz\\Python\\include\\Python.h"
#endif
#endif



// TODO: reference additional headers your program requires here
