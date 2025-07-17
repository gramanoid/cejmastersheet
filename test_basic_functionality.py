#!/usr/bin/env python3
"""
Basic functionality test for the Excel Transformer application.
Tests that the modules can be imported and basic functions work without errors.
"""
import sys
import logging

def test_imports():
    """Test that all modules can be imported successfully."""
    print("Testing imports...")
    try:
        import config
        print("✓ config module imported successfully")
    except Exception as e:
        print(f"✗ Failed to import config: {e}")
        return False
    
    try:
        import excel_transformer
        print("✓ excel_transformer module imported successfully")
    except Exception as e:
        print(f"✗ Failed to import excel_transformer: {e}")
        return False
    
    try:
        import streamlit_app
        print("✓ streamlit_app module imported successfully")
    except Exception as e:
        print(f"✗ Failed to import streamlit_app: {e}")
        return False
    
    return True

def test_config_values():
    """Test that config values are properly defined."""
    print("\nTesting config values...")
    import config
    
    required_attrs = [
        'APP_NAME', 'VERSION', 'DUAL_LANG_INPUT_SHEET_NAME',
        'SINGLE_LANG_INPUT_SHEET_NAME', 'OUTPUT_FILE_BASENAME',
        'PLATFORM_NAMES', 'OUTPUT_COLUMNS_BASE'
    ]
    
    for attr in required_attrs:
        if hasattr(config, attr):
            value = getattr(config, attr)
            print(f"✓ {attr} = {value}")
        else:
            print(f"✗ Missing required config attribute: {attr}")
            return False
    
    return True

def test_safe_to_numeric():
    """Test the safe_to_numeric function with various inputs."""
    print("\nTesting safe_to_numeric function...")
    import pandas as pd
    from excel_transformer import safe_to_numeric
    
    test_cases = [
        (5, 5, "numeric value"),
        ("5", 5, "string number"),
        ("", 0, "empty string"),
        (None, 0, "None value"),
        (pd.NA, 0, "pandas NA"),
        ("abc", 0, "non-numeric string"),
        (5.5, 5.5, "float value"),
    ]
    
    all_passed = True
    for value, expected, description in test_cases:
        result = safe_to_numeric(value, 0, "test_column")
        if result == expected:
            print(f"✓ {description}: {value} → {result}")
        else:
            print(f"✗ {description}: {value} → {result} (expected {expected})")
            all_passed = False
    
    return all_passed

def test_platform_names():
    """Test that platform names are properly configured."""
    print("\nTesting platform configurations...")
    from config import PLATFORM_NAMES
    
    expected_platforms = ["youtube", "meta", "tiktok", "programmatic", "audio", "gaming", "amazon"]
    
    for platform in expected_platforms:
        if platform in PLATFORM_NAMES:
            print(f"✓ Platform '{platform}' configured as '{PLATFORM_NAMES[platform]}'")
        else:
            print(f"✗ Missing platform configuration: {platform}")
            return False
    
    return True

def main():
    """Run all tests."""
    print("=== Excel Transformer Basic Functionality Test ===\n")
    
    # Suppress logging output during tests
    logging.getLogger().setLevel(logging.CRITICAL)
    
    tests = [
        test_imports,
        test_config_values,
        test_safe_to_numeric,
        test_platform_names,
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            if test():
                passed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"✗ Test {test.__name__} raised exception: {e}")
            failed += 1
    
    print(f"\n=== Test Summary ===")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    print(f"Total: {passed + failed}")
    
    return failed == 0

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)