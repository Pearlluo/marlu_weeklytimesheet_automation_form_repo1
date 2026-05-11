from flask import Flask, jsonify, render_template, request

from GetData1 import (
    get_workers_for_app,
    handle_sign_out,
    handle_sign_in,
    find_latest_today_timesheet_record,
)

app = Flask(__name__)

CACHE = {
    "workers": None,
    "supervisors": None
}


def load_cache():
    print("Loading workers and supervisors from SharePoint...")

    workers = get_workers_for_app()

    excluded_names = {
        "jameson joel",
        "odriscoll ciara",
        "shananhan james",
        "kalma guy",
    }

    supervisors = [
        {
            "name": w.get("name", ""),
            "opms": w.get("opms", ""),
            "position": w.get("position", ""),
            "site": w.get("site", ""),
            "project": w.get("project", ""),
        }
        for w in workers
        if w.get("is_supervisor_candidate")
           and w.get("name", "").strip().lower() not in excluded_names
    ]

    CACHE["workers"] = workers
    CACHE["supervisors"] = supervisors

    print(f"Loaded workers: {len(workers)}")
    print(f"Loaded supervisors: {len(supervisors)}")


@app.route("/")
def index():
    if CACHE["workers"] is None:
        load_cache()

    return render_template("form1.html")


@app.route("/api/workers", methods=["GET"])
def api_workers():
    try:
        if CACHE["workers"] is None:
            load_cache()

        return jsonify(CACHE["workers"])
    except Exception as e:
        print("Error in /api/workers:", e)
        return jsonify({"error": str(e)}), 500


@app.route("/api/supervisors", methods=["GET"])
def api_supervisors():
    try:
        if CACHE["supervisors"] is None:
            load_cache()

        return jsonify(CACHE["supervisors"])
    except Exception as e:
        print("Error in /api/supervisors:", e)
        return jsonify({"error": str(e)}), 500


@app.route("/api/refresh-cache", methods=["POST"])
def api_refresh_cache():
    try:
        load_cache()
        return jsonify({"message": "Cache refreshed successfully."})
    except Exception as e:
        print("Error in /api/refresh-cache:", e)
        return jsonify({"error": str(e)}), 500


@app.route("/api/status/<opms>", methods=["GET"])
def api_status(opms):
    try:
        latest = find_latest_today_timesheet_record(opms)

        if latest is None:
            return jsonify({
                "current_status": "No record today",
                "next_action": "Sign Out",
                "latest": None
            })

        latest_status = str(latest.get("status", "")).strip()

        if latest_status.lower() == "sign out":
            next_action = "Sign In"
        elif latest_status.lower() == "sign in":
            next_action = "Sign Out"
        else:
            next_action = "Unknown"

        return jsonify({
            "current_status": latest_status,
            "next_action": next_action,
            "latest": latest
        })

    except Exception as e:
        print("Error in /api/status:", e)
        return jsonify({"error": str(e)}), 400


def validate_payload(data):
    if not data:
        raise Exception("No data received.")

    required = ["name", "opms", "supervisor", "reason"]

    missing = [x for x in required if not str(data.get(x, "")).strip()]
    if missing:
        raise Exception(f"Missing required field(s): {', '.join(missing)}")


@app.route("/api/sign-out", methods=["POST"])
def api_sign_out():
    try:
        data = request.get_json()
        print("Sign Out payload:", data)

        validate_payload(data)

        result = handle_sign_out(data)

        return jsonify({
            "message": "Sign Out recorded successfully.",
            "result": result,
        })

    except Exception as e:
        print("Error in /api/sign-out:", e)
        return jsonify({"error": str(e)}), 400


@app.route("/api/sign-in", methods=["POST"])
def api_sign_in():
    try:
        data = request.get_json()
        print("Sign In payload:", data)

        validate_payload(data)

        result = handle_sign_in(data)

        return jsonify({
            "message": "Sign In recorded successfully.",
            "result": result,
        })

    except Exception as e:
        print("Error in /api/sign-in:", e)
        return jsonify({"error": str(e)}), 400


if __name__ == "__main__":
    load_cache()
    app.run(debug=True, port=5000)