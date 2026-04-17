import sys

import win32com.client


SW_DOC_PART = 1
SW_SEL_FACE = 2


def main() -> int:
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        print("ERROR=Could not connect to the running SOLIDWORKS instance.")
        return 1

    model = sw.ActiveDoc
    if model is None:
        print("ERROR=SOLIDWORKS has no active document.")
        return 2

    if model.GetType != SW_DOC_PART:
        print("ERROR=Active document is not a part document.")
        print(f"DOC_TYPE={model.GetType}")
        return 3

    sel_mgr = model.SelectionManager
    sel_count = sel_mgr.GetSelectedObjectCount2(-1)
    if sel_count < 1:
        print("ERROR=No selection found. Select one face in SOLIDWORKS first.")
        return 4

    sel_type = sel_mgr.GetSelectedObjectType3(1, -1)
    if sel_type != SW_SEL_FACE:
        print("ERROR=The first selected object is not a face.")
        print(f"SELECTED_TYPE={sel_type}")
        return 5

    face = sel_mgr.GetSelectedObject6(1, -1)
    area_m2 = float(face.GetArea())

    print(f"TITLE={model.GetTitle}")
    print(f"SELECTED_COUNT={sel_count}")
    print(f"FACE_AREA_M2={area_m2:.15g}")
    print(f"FACE_AREA_MM2={area_m2 * 1_000_000:.15g}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
