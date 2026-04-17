import argparse
import sys

import win32com.client


SW_DOC_PART = 1
SW_SEL_FACE = 2


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--velocity", type=float, required=True, help="Air velocity in m/s.")
    parser.add_argument(
        "--density",
        type=float,
        default=1.2,
        help="Air density in kg/m^3. Defaults to 1.2.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

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
        print("ERROR=No selection found. Select one inlet face in SOLIDWORKS first.")
        return 4

    sel_type = sel_mgr.GetSelectedObjectType3(1, -1)
    if sel_type != SW_SEL_FACE:
        print("ERROR=The first selected object is not a face.")
        print(f"SELECTED_TYPE={sel_type}")
        return 5

    face = sel_mgr.GetSelectedObject6(1, -1)
    area_m2 = float(face.GetArea())
    flow_m3s = area_m2 * args.velocity
    flow_ls = flow_m3s * 1000.0
    flow_m3h = flow_m3s * 3600.0
    velocity_pressure_pa = 0.5 * args.density * args.velocity * args.velocity

    print(f"TITLE={model.GetTitle}")
    print(f"FACE_AREA_M2={area_m2:.15g}")
    print(f"FACE_AREA_MM2={area_m2 * 1_000_000:.15g}")
    print(f"VELOCITY_MPS={args.velocity:.15g}")
    print(f"DENSITY_KG_M3={args.density:.15g}")
    print(f"FLOW_M3_S={flow_m3s:.15g}")
    print(f"FLOW_L_S={flow_ls:.15g}")
    print(f"FLOW_M3_H={flow_m3h:.15g}")
    print(f"VELOCITY_PRESSURE_PA={velocity_pressure_pa:.15g}")
    print("STATIC_PRESSURE_NOTE=Static pressure cannot be uniquely determined from area and velocity alone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
