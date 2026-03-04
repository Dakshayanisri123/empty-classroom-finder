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

        if not timetable_file or not rooms_file:
            return render_template("index.html", message="Upload both Excel files.")

        if not selected_day or len(selected_hours) == 0:
            return render_template("index.html",
                                   message="Select one day and at least 1 hour.",
                                   selected_day=selected_day,
                                   selected_hours=selected_hours)

        try:
            timetable_df = pd.read_excel(timetable_file, header=1)
            timetable_df.columns = timetable_df.columns.astype(str).str.strip()

            # Find correct sheet in rooms file
            xls = pd.ExcelFile(rooms_file)
            rooms_df = None

            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet)
                temp_df.columns = temp_df.columns.astype(str).str.strip()
                if "Room No" in temp_df.columns:
                    rooms_df = temp_df
                    break

            if rooms_df is None:
                return render_template("index.html", message="Correct sheet not found in rooms file.")

        except Exception as e:
            return render_template("index.html", message=f"Excel read error: {e}")

        columns_to_check = [f"{selected_day}{h}" for h in selected_hours]

        # -------- NEW LOGIC --------
        room_groups = {}

        for _, row in timetable_df.iterrows():

            room_no = str(row.iloc[0]).strip()

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
                    return render_template("index.html", message=f"{col} not found.")

                for r in rows:
                    cell = r[col]

                    if not (pd.isna(cell) or str(cell).strip() == "-"):
                        room_free = False
                        break

                if not room_free:
                    break

            if room_free:
                empty_rooms.append(base_room)

        # -------- GET ROOM DETAILS --------
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

    return render_template("index.html",
                           rooms=rooms,
                           message=message,
                           selected_day=selected_day,
                           selected_hours=selected_hours)


if __name__ == "__main__":
    app.run(debug=True)