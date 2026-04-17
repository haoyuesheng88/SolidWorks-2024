import sys

import win32com.client


SW_DOC_PART = 1
SW_SOLID_BODY = 0


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

    bodies = model.GetBodies2(SW_SOLID_BODY, True)
    body_count = len(bodies) if bodies else 0
    face_count = sum(body.GetFaceCount() for body in bodies) if bodies else 0

    mass = model.Extension.CreateMassProperty()

    print(f"TITLE={model.GetTitle}")
    print(f"BODY_COUNT={body_count}")
    print(f"FACE_COUNT={face_count}")

    if mass is not None:
        print(f"VOLUME_M3={mass.Volume:.15g}")
        print(f"VOLUME_MM3={mass.Volume * 1_000_000_000:.15g}")
        print(f"SURFACE_AREA_M2={mass.SurfaceArea:.15g}")
        print(f"SURFACE_AREA_MM2={mass.SurfaceArea * 1_000_000:.15g}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
