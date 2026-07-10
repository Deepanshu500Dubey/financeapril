#!/usr/bin/env python
"""Test if forecast_editor imports properly"""

print("Testing import...")

try:
    from forecast_editor import get_global_editor, reset_global_editor, ForecastEditor
    print("✅ Successfully imported from forecast_editor")
    
    # Try to create an instance
    editor = get_global_editor()
    print(f"✅ get_global_editor() returned: {type(editor)}")
    
    # Try to use it
    changes = editor.parse_and_apply("increase revenue by 8%")
    print(f"✅ parse_and_apply returned {len(changes)} changes")
    
except ImportError as e:
    print(f"❌ ImportError: {e}")
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()