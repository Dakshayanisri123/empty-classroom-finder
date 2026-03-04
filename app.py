from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def home():
    rooms = []
    message = ""
    selected_day = None
    selected_hours = []

    if request.method == "POST":

        timetable_file = request.files.get("timetable")
        rooms_file = request.files.get("rooms")
        selected_day = request.form.get("day")
        selected_hours = request.form.getlist("hour")

        if not timetable_file or timetable_file.filename == "":
            return render_template("index.html", message="Please upload timetable file")

        if not rooms_file or rooms_file.filename == "":
            return render_template("index.html", message="Please upload rooms master file")

        if not selected_day or len(selected_hours) == 0:
            return render_template(
                "index.html",
                message="Select day and at least one hour",
                selected_day=selected_day,
                selected_hours=selected_hours
            )

        try:
            # MEMORY OPTIMIZED EXCEL READ
            timetable_df = pd.read_excel(
                timetable_file,
                header=1,
                engine="openpyxl"
            )

            timetable_df.fillna("-", inplace=True)
            timetable_df.columns = timetable_df.columns.astype(str).str.strip()

            # Read rooms master file
            rooms_df = pd.read_excel(rooms_file, engine="openpyxl")
            rooms_df.columns = rooms_df.columns.astype(str).str.strip()

        except Exception as e:
            return render_template("index.html", message=f"Excel read error: {e}")

        columns_to_check = [f"{selected_day}{h}" for h in selected_hours]

        # GROUP ROOMS (C007-A, C007-B etc → C007)
        room_groups = {}

        for row in timetable_df.itertuples(index=False):

            room_no = str(row[0]).strip()

            if room_no == "-" or room_no == "nan":
                continue

            base_room = room_no.split("-")[0]

            if base_room not in room_groups:
                room_groups[base_room] = []

            room_groups[base_room].append(row)

        empty_rooms = []

        for base_room, rows in room_groups.items():

            room_free = True

            for col in columns_to_check:

                if col not in timetable_df.columns:
                    return render_template("index.html", message=f"{col} not found")

                col_index = timetable_df.columns.get_loc(col)

                for r in rows:
                    cell = r[col_index]

                    if str(cell).strip() != "-":
                        room_free = False
                        break

                if not room_free:
                    break

            if room_free:
                empty_rooms.append(base_room)

        # GET ROOM DETAILS
        for base_room in empty_rooms:

            match = rooms_df[
                rooms_df["Room No"].astype(str).str.strip() == base_room
            ]

            if not match.empty:
                rooms.append({
                    "room": base_room,
                    "type": match.iloc[0]["CR/LAB"],
                    "strength": match.iloc[0]["TOTAL"]
                })

        if not rooms:
            message = "No empty rooms available for selected hours."

    return render_template(
        "index.html",
        rooms=rooms,
        message=message,
        selected_day=selected_day,
        selected_hours=selected_hours
    )


if __name__ == "__main__":
    app.run()