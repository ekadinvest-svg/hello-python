"""
סקריפט לבניית גרסאות חדשות של האפליקציה
"""
import subprocess
import sys
from pathlib import Path
from datetime import datetime

# ייבוא מידע גרסה
sys.path.insert(0, str(Path(__file__).parent / "src"))
from version import __version__, __app_name__


def build_exe():
    """בניית קובץ EXE"""
    print(f"🔨 בונה {__app_name__} גרסה {__version__}")
    print(f"📅 תאריך: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 50)
    
    # ניקוי build קודם
    print("🧹 מנקה קבצים ישנים...")
    build_dir = Path("build")
    dist_dir = Path("dist")
    
    if build_dir.exists():
        import shutil
        shutil.rmtree(build_dir)
    
    # בניית ה-EXE
    print("⚙️  בונה את האפליקציה...")
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        "workout_tracker.spec"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ הבנייה הושלמה בהצלחה!")
        print(f"📦 הקובץ נמצא ב: {dist_dir / 'מעקב_אימונים.exe'}")
        
        # הצגת גודל הקובץ
        exe_path = dist_dir / "מעקב_אימונים.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📏 גודל הקובץ: {size_mb:.1f} MB")
    else:
        print("❌ הבנייה נכשלה!")
        print(result.stderr)
        return False
    
    return True


def main():
    """פונקציה ראשית"""
    print("=" * 50)
    print(f"   🏋️ בניית {__app_name__}")
    print("=" * 50)
    print()
    
    if build_exe():
        print()
        print("=" * 50)
        print("   🎉 הבנייה הושלמה בהצלחה!")
        print("=" * 50)
        print()
        print("💡 טיפים:")
        print("   1. הקובץ נמצא בתיקיית dist/")
        print("   2. ניתן להעתיק אותו לכל מחשב")
        print("   3. לא נדרש Python מותקן")
        print("   4. הנתונים לא נכללו (ייווצרו בהרצה)")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
