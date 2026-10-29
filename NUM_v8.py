import math
import pythoncom
import win32com.client
from win32com.client import VARIANT

# ==================================================
# USTAWIENIA
# ==================================================

TOL = 0.01
POINT_LAYER = "NUM_POINTS"

# ==================================================
# INPUT
# ==================================================

def ask_string(prompt, default):

    value = input(f"{prompt} [{default}]: ").strip()

    if value == "":
        return default

    return value


def ask_float(prompt, default):

    value = input(f"{prompt} [{default}]: ").strip()

    if value == "":
        return default

    try:
        return float(value)
    except:
        return default


def ask_int(prompt, default):

    value = input(f"{prompt} [{default}]: ").strip()

    if value == "":
        return default

    try:
        return int(value)
    except:
        return default


# ==================================================
# WARSTWA NUM_POINTS
# ==================================================

def get_or_create_layer(doc, layer_name):

    try:
        return doc.Layers.Item(layer_name)
    except:
        return doc.Layers.Add(layer_name)


def get_existing_numbered_points(modelspace):

    points = []

    for obj in modelspace:

        try:

            if (
                obj.ObjectName == "AcDbPoint"
                and obj.Layer == POINT_LAYER
            ):

                pt = obj.Coordinates

                points.append(
                    (
                        pt[0],
                        pt[1]
                    )
                )

        except:
            pass

    return points


def point_is_already_numbered(
    point,
    existing_points,
    tolerance
):

    for pt in existing_points:

        if math.hypot(
            point[0] - pt[0],
            point[1] - pt[1]
        ) <= tolerance:

            return True

    return False


def create_marker_point(
    modelspace,
    point
):

    marker = modelspace.AddPoint(
        VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (
                point[0],
                point[1],
                0.0
            )
        )
    )

    marker.Layer = POINT_LAYER


# ==================================================
# ROTACJA
# ==================================================

def normalize_rotation(angle):

    while angle < 0:
        angle += 2 * math.pi

    while angle >= 2 * math.pi:
        angle -= 2 * math.pi

    if math.pi / 2 < angle < 3 * math.pi / 2:
        angle += math.pi

    while angle >= 2 * math.pi:
        angle -= 2 * math.pi

    return angle


# ==================================================
# KÄ„T WIERZCHOĹKA
# ==================================================

def get_vertex_angle(coords, idx):

    vert_count = len(coords) // 2

    if vert_count < 2:
        return 0.0

    if idx == vert_count - 1:

        x1 = coords[(idx - 1) * 2]
        y1 = coords[(idx - 1) * 2 + 1]

        x2 = coords[idx * 2]
        y2 = coords[idx * 2 + 1]

    else:

        x1 = coords[idx * 2]
        y1 = coords[idx * 2 + 1]

        x2 = coords[(idx + 1) * 2]
        y2 = coords[(idx + 1) * 2 + 1]

    return math.atan2(
        y2 - y1,
        x2 - x1
    )


# ==================================================
# ZBIERANIE PUNKTĂ“W
# ==================================================

def collect_points(selection_set):

    points = []

    for obj in selection_set:

        if obj.ObjectName != "AcDbPolyline":
            continue

        coords = list(obj.Coordinates)

        vertex_count = len(coords) // 2

        for i in range(vertex_count):

            x = coords[i * 2]
            y = coords[i * 2 + 1]

            angle = get_vertex_angle(coords, i)

            exists = False

            for p in points:

                dx = x - p[0]
                dy = y - p[1]

                if math.hypot(dx, dy) <= TOL:

                    exists = True
                    break

            if not exists:

                points.append(
                    (x, y, angle)
                )

    return points


# ==================================================
# SORTOWANIE
# ==================================================

def sort_points(points):

    return sorted(
        points,
        key=lambda p: (
            round(p[0] / TOL),
            -p[1]
        )
    )


# ==================================================
# FILTR TOLERANCJI
# ==================================================

def apply_tolerance(points, tolerance):

    if tolerance <= 0:
        return points

    result = []

    last_kept = None

    for pt in points:

        if last_kept is None:

            result.append(pt)
            last_kept = pt
            continue

        dist = math.hypot(
            pt[0] - last_kept[0],
            pt[1] - last_kept[1]
        )

        if dist >= tolerance:

            result.append(pt)
            last_kept = pt

    return result


# ==================================================
# TWORZENIE TEKSTU
# ==================================================

def add_text(
    modelspace,
    point,
    base_angle,
    text_string,
    text_height,
    offset_factor,
    side,
    orientation
):

    offset = text_height * offset_factor

    if math.cos(base_angle) < 0:
        base_angle += math.pi

    if side.upper() == "L":
        side_angle = base_angle + math.pi / 2
    else:
        side_angle = base_angle - math.pi / 2

    text_x = point[0] + math.cos(side_angle) * offset
    text_y = point[1] + math.sin(side_angle) * offset

    insertion_point = VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (text_x, text_y, 0.0)
    )

    txt = modelspace.AddText(
        text_string,
        insertion_point,
        text_height
    )

    orientation = orientation.upper()

    if orientation == "POLILINIA":
        rotation = base_angle

    elif orientation == "90":
        rotation = base_angle + math.pi / 2

    else:
        rotation = 0.0

    rotation = normalize_rotation(rotation)

    try:

        txt.Alignment = 10
        txt.TextAlignmentPoint = insertion_point
        txt.Rotation = rotation
        txt.Update()

    except:
        pass

    return txt


# ==================================================
# MAIN
# ==================================================

def main():

    print()

    start_number = ask_int(
        "Numer poczatkowy",
        1
    )

    prefix = ask_string(
        "Przedrostek",
        "P-"
    )

    text_height = ask_float(
        "Wysokosc tekstu",
        2.5
    )

    offset_factor = ask_float(
        "Mnoznik odsuniecia",
        2.0
    )

    side = ask_string(
        "Strona (L/P)",
        "L"
    ).upper()

    orientation = ask_string(
        "Orientacja (Polilinia/90/Poziomo)",
        "90"
    )

    tolerance = ask_float(
        "Tolerancja",
        0.0
    )

    skip_existing = ask_string(
        "Pomijaj juz opisane punkty (T/N)",
        "T"
    ).upper()

    acad = win32com.client.GetActiveObject(
        "AutoCAD.Application"
    )

    doc = acad.ActiveDocument

    get_or_create_layer(
        doc,
        POINT_LAYER
    )

    try:

        ss = doc.PickfirstSelectionSet

    except:

        print("Brak zaznaczenia.")
        return

    points = collect_points(ss)

    print()
    print(f"Wybrane obiekty: {ss.Count}")
    print(f"Unikalne punkty: {len(points)}")

    points = sort_points(points)
    points = apply_tolerance(points, tolerance)

    print(
        f"Punkty po tolerancji: {len(points)}"
    )

    modelspace = doc.ModelSpace

    existing_numbered_points = []

    if skip_existing == "T":

        existing_numbered_points = (
            get_existing_numbered_points(
                modelspace
            )
        )

        print(
            f"Znaleziono "
            f"{len(existing_numbered_points)} "
            f"wczesniej oznaczonych punktow"
        )

    current_number = start_number

    print()
    print("Tworzenie numerow...")

    for pt in points:

        if skip_existing == "T":

            if point_is_already_numbered(
                pt,
                existing_numbered_points,
                max(tolerance, TOL)
            ):
                continue

        add_text(
            modelspace,
            (pt[0], pt[1]),
            pt[2],
            f"{prefix}{current_number}",
            text_height,
            offset_factor,
            side,
            orientation
        )

        create_marker_point(
            modelspace,
            pt
        )

        existing_numbered_points.append(
            (
                pt[0],
                pt[1]
            )
        )

        current_number += 1

    doc.Regen(1)

    print()
    print(
        f"Ostatni numer: {current_number - 1}"
    )


if __name__ == "__main__":
    main()
