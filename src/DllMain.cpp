// MermaidPreview - EmEditor Plugin
// Renders Mermaid diagrams in a WebView2 sidebar panel.

// Windows headers must be included before plugin.h / etlframe.h
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>

// Define the frame class name BEFORE including etlframe.h.
// etlframe.h uses this to forward-declare the class and typedef CETLFrameX.
#define ETL_FRAME_CLASS_NAME CMermaidFrame

// Include the full ETL framework.
// This provides: DllMain, OnCommand, QueryStatus, OnEvents,
// GetMenuTextID, GetStatusMessageID, GetBitmapID, PlugInProc
// All export functions are defined here (no EE_EXTERN_ONLY).
#include "etlframe.h"

// Include complete types for unique_ptr members
#include "WebView2Manager.h"
#include "BunRenderer.h"

// Include the frame class definition.
// CETLFrame<T> template is now fully defined by etlframe.h,
// so CMermaidFrame can inherit from it.
#include "MermaidPreview.h"

// Generate the factory functions for creating/deleting frame instances.
// This macro expands to:
//   CETLFrameX* _ETLCreateFrame() { return new CMermaidFrame; }
//   void _ETLDeleteFrame(CETLFrameX* p) { delete static_cast<CMermaidFrame*>(p); }
_ETL_IMPLEMENT
