"""全テストケースをテンプレートエンジンv2で実行"""
import json, sys, os, glob
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# exec pattern to avoid Windows encoding issues
engine_path = os.path.join(os.path.dirname(__file__), '..', 'backend', 'template_engine_v2.py')
exec(open(engine_path, encoding='utf-8').read())

test_dir = os.path.dirname(__file__)
test_files = sorted(glob.glob(os.path.join(test_dir, 'test_*.json')))

total = 0
passed = 0
failed = 0
errors = []

for f in test_files:
    name = os.path.basename(f)
    print(f"\n{'='*60}")
    print(f"  TEST: {name}")
    print(f"{'='*60}")

    with open(f, 'r', encoding='utf-8') as fh:
        data = json.load(fh)

    groups = data.get('processing_groups', [])
    test_passed = True

    for g in groups:
        total += 1
        try:
            result = render_group(g)
            if '変数不足' in result:
                print(f"  [WARN] Group '{g.get('name','')}': 変数不足あり")
                test_passed = False
                failed += 1
                errors.append(f"{name} / {g.get('name','')}: 変数不足")
            else:
                passed += 1
                # Show first 200 chars
                preview = result[:200].replace('\n', ' | ')
                print(f"  [OK] {g.get('name','')}: {preview}...")
        except Exception as e:
            failed += 1
            test_passed = False
            errors.append(f"{name} / {g.get('name','')}: {e}")
            print(f"  [FAIL] {g.get('name','')}: {e}")

print(f"\n{'='*60}")
print(f"  RESULTS: {passed}/{total} passed, {failed} failed")
print(f"{'='*60}")
if errors:
    print("\nErrors:")
    for e in errors:
        print(f"  - {e}")
