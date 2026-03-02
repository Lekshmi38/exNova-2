#complete class 


from collections import defaultdict
import random

# -------------------------
# INPUT DATA
# -------------------------

class_group = { 
    "S7CS1": "S7CS", "S7CS2": "S7CS",
    "S7IT": "S7CS",
    "S7EC": "S7ECER", "S7ER": "S7ECER",
    "S7CE": "S7ECER",
    "S5CS1": "S5CS", "S5CS2": "S5CS",
    "S5IT": "S5CS",
    "S5EC": "S5ECE", "S5ER": "S5ECE", "S5CE": "S5ECE",
    "S3CS1": "S3CS", "S3CS2": "S3CS", "S3CS3": "S3CS",
    "S3IT": "S3CS",
    "S3EC": "S3ECER", "S3ER": "S3ECER", "S3CE": "S3ECER"   
}

room_capacity_18 = {
    "Room1":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room2":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room3":  {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
    "Room4":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room5":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room6":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room7":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room8":  {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
    "Room9":  {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room10": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room11": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room12": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room13": {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
    "Room14": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room15": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room16": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room17": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room18": {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
    "Room19": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room20": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room21": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room22": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room23": {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
    "Room24": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room25": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room26": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room27": {"Left1":7, "Left2":7, "Left3":7, "Middle1":6, "Middle2":6, "Right1":7, "Right2":7, "Right3":7},
    "Room28": {"Left1":6, "Left2":6, "Left3":6, "Middle1":6, "Middle2":6, "Right1":6, "Right2":6, "Right3":6},
}

block_order = list(room_capacity_18["Room1"].keys())

classes = {
    "S3CS1": 70, "S3CS2": 71, "S3CS3": 71, "S3IT": 66, "S3EC": 68, "S3ER": 66, "S3CE": 65,
    "S5CS1": 71, "S5CS2": 72, "S5IT": 61, "S5CE": 56, "S5EC": 69, "S5ER": 56,
    "S7CS1": 67, "S7CS2": 67, "S7IT": 60, "S7EC": 64, "S7ER": 63, "S7CE": 62
}

timetable_data = {
    "22-08-25": {
        "FN": {
            "S3CS1": "Digital Electronics and Design", "S3CS2": "Digital Electronics and Design", "S3CS3": "Digital Electronics and Design",
            "S3IT": "Digital Electronics and Design", "S3EC": "Maths for EC", "S3CE": "Maths for PS", "S3ER": "Maths for EC",
            "S5CS1": "Diasaster Managemnt", "S5CS2": "Diasaster Managemnt", "S5IT": "flat",
            "S5EC": "Analog & Digital ", "S5CE": "Structural Analysis", "S5ER": "MicroProcessor",
            "S7CS1": "AI", "S7CS2": "AI",  "S7IT": "ML",
            "S7EC": "OPtical Fibre", "S7CE": "Ground Improve", "S7ER": "Web Programming"
        },
        "AN": {
            "S5CS1": "Computer Networks", "S5CS2": "Computer Networks", "S5IT": "Software Engineering",
            "S5EC": "Control Systems", "S5CE": "Transportation Engineering", "S5ER": "Power Electronics",
            "S7CS1": "Cloud Computing", "S7CS2": "Cloud Computing", "S7IT": "Mobile Computing",
            "S7EC": "Embedded Systems", "S7CE": "Environmental Engineering", "S7ER": "Renewable Energy"
        }
    }
}

# -------------------------
# HELPER FUNCTIONS
# -------------------------
def get_branch(cls):
    """
    Extract branch from class code
    Examples:
    S3CS1 -> CS
    S5IT  -> IT
    S7EC  -> EC
    S7ER  -> ER
    S7CE  -> CE
    """
    return cls[2:4]


def can_place(
    student_class,
    student_subject,
    blocks_in_room,
    block_index,
    block_order,
    room_no,
    max_classes_strict=3
):
    current_sem = int(student_class[1])
    current_group = class_group.get(student_class, student_class)
    current_branch = get_branch(student_class)

    block_name = block_order[block_index]

    # --------------------------------------------------
    # 1. BLOCK-LEVEL: ONLY ONE CLASS PER BLOCK
    # --------------------------------------------------
    if blocks_in_room[block_name]:
        if student_class not in blocks_in_room[block_name]:
            return False

    # --------------------------------------------------
    # 2. ROOM-LEVEL: MAX 3 CLASSES (ROOM LOCKING)
    # --------------------------------------------------
    room_classes = set()
    for block in blocks_in_room.values():
        room_classes.update(block.keys())

    if student_class not in room_classes and len(room_classes) >= max_classes_strict:
        return False

    # --------------------------------------------------
    # 3. ADJACENCY RULES
    # --------------------------------------------------
    def violates(adj_block):
        for cls in blocks_in_room[adj_block]:
            adj_sem = int(cls[1])
            adj_group = class_group.get(cls, cls)
            adj_branch = get_branch(cls)

            # ❌ same semester adjacency (S5–S5 etc.)
            if adj_sem == current_sem:
                return True

            # ❌ same group adjacency
            if adj_group == current_group:
                return True

            # ❌ CS next to CS (any semester)
            if adj_branch == "CS" and current_branch == "CS":
                return True

        return False

    # Left adjacency
    if block_index > 0:
        if violates(block_order[block_index - 1]):
            return False

    # Right adjacency
    if block_index < len(block_order) - 1:
        if violates(block_order[block_index + 1]):
            return False

    return True

# -------------------------
# SESSION ALLOCATION
# -------------------------
def allocate_session(session_name):
    remaining = classes.copy()
    classrooms = {}
    class_assigned_count = defaultdict(int)

    available_subjects = timetable_data['22-08-25'].get(session_name, {})

    active_classes = [
        cls for cls in classes
        if cls in available_subjects and available_subjects[cls].strip()
    ]

    s7_open_counts = {}

    

    for room_no in range(1, len(room_capacity_18) + 1):
        room_name = f"Room{room_no}"
        blocks_in_room = {block: {} for block in block_order}
        capacity_map = room_capacity_18[room_name].copy()

        for block in block_order:
            block_index = block_order.index(block)

            while capacity_map[block] > 0:
                room_classes = set()
                for b in blocks_in_room.values():
                    room_classes.update(b.keys())

                possible_classes = [
                    cls for cls in active_classes
                    if s7_open_counts.get(cls, remaining[cls]) > 0
                ]

                if not possible_classes:
                    break

                # --------------------------------------------------
                # ROOM-AWARE PRIORITY (KEY OPTIMIZATION)
                # --------------------------------------------------
                def class_priority(c):
                    count = s7_open_counts.get(c, remaining[c])
                    if c in room_classes:
                        return (3, count)   # continue same class
                    return (2, count)       # new class (if allowed)

                possible_classes.sort(key=class_priority, reverse=True)

                placed = False
                for student_class in possible_classes:
                    subject = available_subjects.get(student_class, "")

                    if can_place(
                        student_class,
                        subject,
                        blocks_in_room,
                        block_index,
                        block_order,
                        room_no
                    ):
                        assign = min(
                            s7_open_counts.get(student_class, remaining[student_class]),
                            capacity_map[block]
                        )

                        blocks_in_room[block][student_class] = {
                            "count": assign,
                            "subject": subject
                        }

                        if student_class in s7_open_counts:
                            s7_open_counts[student_class] -= assign
                        else:
                            remaining[student_class] -= assign

                        capacity_map[block] -= assign
                        class_assigned_count[student_class] += assign
                        placed = True
                        break

                if not placed:
                    break

        classrooms[room_name] = blocks_in_room

    leftovers = {
        cls: s7_open_counts.get(cls, remaining[cls])
        for cls in active_classes
    }

    return classrooms, leftovers, class_assigned_count, []

# -------------------------
# STACK LEFTOVERS
# -------------------------
def stack_leftovers(classrooms, leftovers, session_name):
    available_subjects = timetable_data['22-08-25'].get(session_name, {})

    for room_name in sorted(classrooms, key=lambda x: int(x.replace("Room", ""))):
        room_no = int(room_name.replace("Room", ""))
        blocks = classrooms[room_name]

        for block in block_order:
            block_dict = blocks[block]
            capacity_left = room_capacity_18[room_name][block] - sum(
                info["count"] for info in block_dict.values()
            )

            if capacity_left <= 0:
                continue

            block_index = block_order.index(block)

            for cls in list(leftovers):
                if leftovers[cls] <= 0:
                    continue

                subject = available_subjects.get(cls, "")
                if not subject:
                    continue

                if can_place(
                    cls,
                    subject,
                    blocks,
                    block_index,
                    block_order,
                    room_no
                ):
                    assign = min(leftovers[cls], capacity_left)

                    block_dict[cls] = {
                        "count": block_dict.get(cls, {"count": 0})["count"] + assign,
                        "subject": subject
                    }

                    leftovers[cls] -= assign
                    capacity_left -= assign

                if capacity_left <= 0:
                    break

    return classrooms, leftovers

# -------------------------
# CLASS-WISE TABLE & VISUALIZATION
# -------------------------
def generate_classwise_table(class_assigned_count, classrooms):
    classwise_table = defaultdict(list)
    roll_start = {cls:1 for cls in classes}
    for room, blocks in classrooms.items():
        for block, block_dict in blocks.items():
            for cls, info in block_dict.items():
                count = info["count"]
                subject = info["subject"]
                roll_end = roll_start[cls] + count - 1
                classwise_table[cls].append({
                    "Roll Start": roll_start[cls],
                    "Roll End": roll_end,
                    "Room": room,
                    "Block": block,
                    "Subject": subject
                })
                roll_start[cls] = roll_end + 1
    return classwise_table

def visualize_all_rooms(classrooms, file):
    for room_name in sorted(classrooms, key=lambda x: int(x.replace("Room",""))):
        blocks = classrooms[room_name]

        file.write(f"\n=== {room_name} ===\n")
        file.write(f"{'Block':<10} | {'Class':<8} | {'Count':<5} | {'Subject'}\n")
        file.write("-" * 50 + "\n")

        for block in block_order:
            block_dict = blocks[block]
            if not block_dict:
                file.write(f"{block:<10} | {'-':<8} | {'-':<5} | {'-'}\n")
                continue

            for cls, info in block_dict.items():
                file.write(
                    f"{block:<10} | {cls:<8} | {info['count']:<5} | {info['subject']}\n"
                )

        file.write("-" * 50 + "\n")

# -------------------------
# EXECUTION
# -------------------------
fn_classrooms, fn_leftover, fn_class_count, fn_skipped = allocate_session('FN')
an_classrooms, an_leftover, an_class_count, an_skipped = allocate_session('AN')

fn_classrooms, fn_leftover = stack_leftovers(fn_classrooms, fn_leftover, 'FN')
an_classrooms, an_leftover = stack_leftovers(an_classrooms, an_leftover, 'AN')

fn_classwise = generate_classwise_table(fn_class_count, fn_classrooms)
an_classwise = generate_classwise_table(an_class_count, an_classrooms)

# -------------------------
# WRITE OUTPUT TO FILE
# -------------------------
with open("seating_output.txt", "w", encoding="utf-8") as f:

    f.write("--- FN ROOM-WISE ---\n")
    visualize_all_rooms(fn_classrooms, f)
    f.write(f"Leftovers FN: {fn_leftover}\n")
    f.write(f"Skipped Classes (no exam): {fn_skipped}\n\n")

    f.write("--- AN ROOM-WISE ---\n")
    visualize_all_rooms(an_classrooms, f)
    f.write(f"Leftovers AN: {an_leftover}\n")
    f.write(f"Skipped Classes (no exam): {an_skipped}\n\n")

    f.write("--- FN CLASS-WISE ---\n")
    for cls, allocations in fn_classwise.items():
        if allocations:
            f.write(f"Class {cls}:\n")
            for alloc in allocations:
                f.write(
                    f"  Roll {alloc['Roll Start']}-{alloc['Roll End']} "
                    f"-> {alloc['Room']} ({alloc['Subject']})\n"
                )

    f.write("\n--- AN CLASS-WISE ---\n")
    for cls, allocations in an_classwise.items():
        if allocations:
            f.write(f"Class {cls}:\n")
            for alloc in allocations:
                f.write(
                    f"  Roll {alloc['Roll Start']}-{alloc['Roll End']} "
                    f"-> {alloc['Room']} ({alloc['Subject']})\n"
                )

print("✅ Seating arrangement written to seating_output.txt")
