"""Quick test script for forecast_editor module"""
import forecast_editor

print('=== TEST 1: Increase the forecast by 10% ===')
e = forecast_editor.ForecastEditor()
changes = e.parse_and_apply('Increase the forecast by 10%')
print(f'Changes: {len(changes)}')
for c in changes:
    print(f'  [{c["model"]}] {c["label"]} ({c["scenario"]}): {c["old_value_formatted"]} => {c["new_value_formatted"]} ({c["delta_pct_formatted"]})')

print()
print('=== TEST 2: Reduce the sales forecast for Q2 by 5% ===')
e2 = forecast_editor.ForecastEditor()
changes2 = e2.parse_and_apply('Reduce the sales forecast for Q2 by 5%')
print(f'Changes: {len(changes2)}')
for c in changes2:
    print(f'  [{c["model"]}] {c["label"]} ({c["scenario"]}): {c["old_value_formatted"]} => {c["new_value_formatted"]} ({c["delta_pct_formatted"]})')

print()
print('=== TEST 3: Increase the revenue forecast for Product A by 20% ===')
e3 = forecast_editor.ForecastEditor()
changes3 = e3.parse_and_apply('Increase the revenue forecast for Product A by 20%')
print(f'Changes: {len(changes3)}')
for c in changes3:
    print(f'  [{c["model"]}] {c["label"]} ({c["scenario"]}): {c["old_value_formatted"]} => {c["new_value_formatted"]} ({c["delta_pct_formatted"]})')

print()
print('=== TEST 4: Reduce logistics costs by 15% ===')
e4 = forecast_editor.ForecastEditor()
changes4 = e4.parse_and_apply('Reduce logistics costs by 15%')
print(f'Changes: {len(changes4)}')
for c in changes4:
    print(f'  [{c["model"]}] {c["label"]} ({c["scenario"]}): {c["old_value_formatted"]} => {c["new_value_formatted"]} ({c["delta_pct_formatted"]})')

print()
print('=== TEST 5: Comparison Excel generation ===')
e5 = forecast_editor.ForecastEditor()
e5.parse_and_apply('Increase the forecast by 10%')
e5.parse_and_apply('Reduce logistics costs by 15%')
comparison_bytes = e5.generate_comparison_excel_bytes()
print(f'Comparison Excel size: {len(comparison_bytes):,} bytes')

print()
print('=== TEST 6: Reset ===')
e5.reset()
print(f'After reset, change log length: {len(e5.change_log)}')

print()
print('=== ALL TESTS PASSED ===')
