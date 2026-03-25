import io
import json
import os
from copy import deepcopy
from datetime import datetime
from uuid import uuid4

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data_store.json")
APP_PORT = int(os.environ.get("SAP_CR_MANAGER_PORT", "5055"))
APP_DEBUG = (os.environ.get("SAP_CR_MANAGER_DEBUG", "0").strip() == "1")

app.secret_key = os.environ.get("SAP_CR_MANAGER_SECRET_KEY", "sap-cr-manager-dev-secret")

STATUS_META = {
    "development": {
        "label": "Sviluppo",
        "tone": "development",
    },
    "quality": {
        "label": "Quality",
        "tone": "quality",
    },
    "production": {
        "label": "Produzione",
        "tone": "production",
    },
}

KANBAN_ORDER = ["development", "quality", "production"]

DEFAULT_DATA = {
    "clients": [],
}


def deep_copy_default() -> dict:
    return deepcopy(DEFAULT_DATA)


def sanitize_text(raw_value: str) -> str:
    return (raw_value or "").strip()


def new_id() -> str:
    return uuid4().hex[:12]


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def ensure_data_file() -> None:
    if os.path.exists(DATA_FILE):
        return
    save_data(deep_copy_default())


def normalize_status(raw_value: str) -> str:
    candidate = sanitize_text(raw_value).lower()
    if candidate in STATUS_META:
        return candidate
    return "development"


def normalize_cr(raw_cr: dict) -> dict:
    created_at = sanitize_text(raw_cr.get("created_at")) or now_iso()
    updated_at = sanitize_text(raw_cr.get("updated_at")) or created_at
    return {
        "id": sanitize_text(raw_cr.get("id")) or new_id(),
        "cr_key": sanitize_text(raw_cr.get("cr_key")),
        "created_by": sanitize_text(raw_cr.get("created_by")),
        "description": sanitize_text(raw_cr.get("description")),
        "notes": sanitize_text(raw_cr.get("notes")),
        "status": normalize_status(raw_cr.get("status")),
        "created_at": created_at,
        "updated_at": updated_at,
    }


def normalize_project(raw_project: dict) -> dict:
    crs = [normalize_cr(item) for item in raw_project.get("crs", []) if isinstance(item, dict)]
    return {
        "id": sanitize_text(raw_project.get("id")) or new_id(),
        "name": sanitize_text(raw_project.get("name")),
        "crs": crs,
    }


def normalize_client(raw_client: dict) -> dict:
    projects = [normalize_project(item) for item in raw_client.get("projects", []) if isinstance(item, dict)]
    return {
        "id": sanitize_text(raw_client.get("id")) or new_id(),
        "name": sanitize_text(raw_client.get("name")),
        "projects": projects,
    }


def load_data() -> dict:
    ensure_data_file()
    with open(DATA_FILE, "r", encoding="utf-8") as handle:
        loaded = json.load(handle)

    data = deep_copy_default()
    data["clients"] = [normalize_client(item) for item in loaded.get("clients", []) if isinstance(item, dict)]
    return data


def save_data(data: dict) -> None:
    with open(DATA_FILE, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)


def find_client(data: dict, client_id: str) -> dict | None:
    for client in data["clients"]:
        if client.get("id") == client_id:
            return client
    return None


def find_project(client: dict, project_id: str) -> dict | None:
    for project in client.get("projects", []):
        if project.get("id") == project_id:
            return project
    return None


def find_cr(project: dict, cr_id: str) -> dict | None:
    for cr in project.get("crs", []):
        if cr.get("id") == cr_id:
            return cr
    return None


def count_statuses(clients: list[dict]) -> dict:
    counts = {key: 0 for key in STATUS_META}
    for client in clients:
        for project in client.get("projects", []):
            for cr in project.get("crs", []):
                status = normalize_status(cr.get("status"))
                counts[status] += 1
    return counts


def matches_filters(cr: dict, search_text: str, status_filter: str) -> bool:
    if status_filter != "all" and normalize_status(cr.get("status")) != status_filter:
        return False

    if not search_text:
        return True

    haystack = " ".join(
        [
            sanitize_text(cr.get("cr_key")),
            sanitize_text(cr.get("created_by")),
            sanitize_text(cr.get("description")),
            sanitize_text(cr.get("notes")),
        ]
    ).lower()
    return search_text in haystack


def normalize_filter_value(raw_value: str) -> str:
    return sanitize_text(raw_value).lower()


def build_filter_options(data: dict) -> dict:
    client_names = sorted(
        {sanitize_text(client.get("name")) for client in data["clients"] if sanitize_text(client.get("name"))},
        key=str.lower,
    )
    project_names = sorted(
        {
            sanitize_text(project.get("name"))
            for client in data["clients"]
            for project in client.get("projects", [])
            if sanitize_text(project.get("name"))
        },
        key=str.lower,
    )
    return {
        "clients": client_names,
        "projects": project_names,
    }


def build_global_kanban_columns(global_crs: list[dict]) -> list[dict]:
    columns = []
    for status_key in KANBAN_ORDER:
        meta = STATUS_META[status_key]
        columns.append(
            {
                "key": status_key,
                "label": meta["label"],
                "tone": meta["tone"],
                "crs": [cr for cr in global_crs if cr.get("status") == status_key],
            }
        )
    return columns


def build_export_rows(
    search_text: str = "",
    status_filter: str = "all",
    client_filter: str = "all",
    project_filter: str = "all",
) -> list[dict]:
    data = load_data()
    normalized_search = sanitize_text(search_text).lower()
    normalized_status = sanitize_text(status_filter).lower() or "all"
    normalized_client = normalize_filter_value(client_filter) or "all"
    normalized_project = normalize_filter_value(project_filter) or "all"
    if normalized_status != "all" and normalized_status not in STATUS_META:
        normalized_status = "all"

    rows: list[dict] = []
    sorted_clients = sorted(data["clients"], key=lambda item: item.get("name", "").lower())
    for client in sorted_clients:
        client_name = sanitize_text(client.get("name"))
        if normalized_client != "all" and client_name.lower() != normalized_client:
            continue

        sorted_projects = sorted(client.get("projects", []), key=lambda item: item.get("name", "").lower())
        for project in sorted_projects:
            project_name = sanitize_text(project.get("name"))
            if normalized_project != "all" and project_name.lower() != normalized_project:
                continue

            sorted_crs = sorted(
                project.get("crs", []),
                key=lambda item: (item.get("updated_at", ""), item.get("cr_key", "")),
                reverse=True,
            )
            for cr in sorted_crs:
                if not matches_filters(cr, normalized_search, normalized_status):
                    continue

                status = normalize_status(cr.get("status"))
                rows.append(
                    {
                        "client": client_name,
                        "project": project_name,
                        "cr_key": sanitize_text(cr.get("cr_key")),
                        "created_by": sanitize_text(cr.get("created_by")),
                        "description": sanitize_text(cr.get("description")),
                        "notes": sanitize_text(cr.get("notes")),
                        "status": STATUS_META[status]["label"],
                        "created_at": sanitize_text(cr.get("created_at")).replace("T", " "),
                        "updated_at": sanitize_text(cr.get("updated_at")).replace("T", " "),
                    }
                )

    return rows


def autosize_worksheet_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            cell_value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(cell_value))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 48)


def build_excel_workbook(
    search_text: str = "",
    status_filter: str = "all",
    client_filter: str = "all",
    project_filter: str = "all",
) -> io.BytesIO:
    rows = build_export_rows(
        search_text=search_text,
        status_filter=status_filter,
        client_filter=client_filter,
        project_filter=project_filter,
    )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Lista CR"

    headers = [
        "Cliente",
        "Progetto",
        "ID CR",
        "Creata da",
        "Descrizione",
        "Note",
        "Stato",
        "Creata il",
        "Aggiornata il",
    ]
    worksheet.append(headers)

    for cell in worksheet[1]:
        cell.font = cell.font.copy(bold=True)

    for row in rows:
        worksheet.append(
            [
                row["client"],
                row["project"],
                row["cr_key"],
                row["created_by"],
                row["description"],
                row["notes"],
                row["status"],
                row["created_at"],
                row["updated_at"],
            ]
        )

    worksheet.freeze_panes = "A2"
    autosize_worksheet_columns(worksheet)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def build_view_model(
    search_text: str = "",
    status_filter: str = "all",
    client_filter: str = "all",
    project_filter: str = "all",
) -> dict:
    data = load_data()
    normalized_search = sanitize_text(search_text).lower()
    normalized_status = sanitize_text(status_filter).lower() or "all"
    normalized_client = normalize_filter_value(client_filter) or "all"
    normalized_project = normalize_filter_value(project_filter) or "all"
    if normalized_status != "all" and normalized_status not in STATUS_META:
        normalized_status = "all"

    visible_clients: list[dict] = []
    global_crs: list[dict] = []
    total_projects = 0
    visible_projects = 0
    total_crs = 0
    visible_crs = 0
    filter_options = build_filter_options(data)

    sorted_clients = sorted(data["clients"], key=lambda item: item.get("name", "").lower())
    for client in sorted_clients:
        client_name = sanitize_text(client.get("name"))
        if normalized_client != "all" and client_name.lower() != normalized_client:
            total_projects += len(sorted(client.get("projects", []), key=lambda item: item.get("name", "").lower()))
            total_crs += sum(len(project.get("crs", [])) for project in client.get("projects", []))
            continue

        client_projects: list[dict] = []
        sorted_projects = sorted(client.get("projects", []), key=lambda item: item.get("name", "").lower())
        total_projects += len(sorted_projects)

        for project in sorted_projects:
            project_name = sanitize_text(project.get("name"))
            if normalized_project != "all" and project_name.lower() != normalized_project:
                total_crs += len(project.get("crs", []))
                continue

            decorated_crs = []
            sorted_crs = sorted(
                project.get("crs", []),
                key=lambda item: (item.get("updated_at", ""), item.get("cr_key", "")),
                reverse=True,
            )
            total_crs += len(sorted_crs)

            for cr in sorted_crs:
                if not matches_filters(cr, normalized_search, normalized_status):
                    continue

                status = normalize_status(cr.get("status"))
                meta = STATUS_META[status]
                decorated_cr = deepcopy(cr)
                decorated_cr["status_label"] = meta["label"]
                decorated_cr["tone"] = meta["tone"]
                decorated_cr["client_name"] = client_name
                decorated_cr["project_name"] = project_name
                decorated_cr["client_id"] = client["id"]
                decorated_cr["project_id"] = project["id"]
                decorated_crs.append(decorated_cr)
                global_crs.append(decorated_cr)

            if decorated_crs:
                kanban_columns = []
                for status_key in KANBAN_ORDER:
                    meta = STATUS_META[status_key]
                    kanban_columns.append(
                        {
                            "key": status_key,
                            "label": meta["label"],
                            "tone": meta["tone"],
                            "crs": [cr for cr in decorated_crs if cr.get("status") == status_key],
                        }
                    )

                visible_projects += 1
                visible_crs += len(decorated_crs)
                client_projects.append(
                    {
                        "id": project["id"],
                        "name": project["name"],
                        "crs": decorated_crs,
                        "kanban_columns": kanban_columns,
                        "cr_count": len(project.get("crs", [])),
                    }
                )

        if client_projects:
            visible_clients.append(
                {
                    "id": client["id"],
                    "name": client["name"],
                    "projects": client_projects,
                    "project_count": len(client.get("projects", [])),
                    "cr_count": sum(len(project.get("crs", [])) for project in client.get("projects", [])),
                }
            )

    status_counts = count_statuses(data["clients"])

    return {
        "clients": visible_clients,
        "summary": {
            "clients": len(data["clients"]),
            "projects": total_projects,
            "crs": total_crs,
            "visible_clients": len(visible_clients),
            "visible_projects": visible_projects,
            "visible_crs": visible_crs,
            "status_counts": status_counts,
        },
        "filters": {
            "search_text": sanitize_text(search_text),
            "status_filter": normalized_status,
            "client_filter": normalized_client,
            "project_filter": normalized_project,
        },
        "filter_options": filter_options,
        "status_meta": STATUS_META,
        "global_kanban_columns": build_global_kanban_columns(global_crs),
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }


@app.get("/")
def index():
    view_model = build_view_model(
        search_text=request.args.get("q", ""),
        status_filter=request.args.get("status", "all"),
        client_filter=request.args.get("client", "all"),
        project_filter=request.args.get("project", "all"),
    )
    return render_template("index.html", **view_model)


@app.get("/export/excel")
def export_excel():
    search_text = request.args.get("q", "")
    status_filter = request.args.get("status", "all")
    client_filter = request.args.get("client", "all")
    project_filter = request.args.get("project", "all")
    workbook_stream = build_excel_workbook(
        search_text=search_text,
        status_filter=status_filter,
        client_filter=client_filter,
        project_filter=project_filter,
    )
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        workbook_stream,
        as_attachment=True,
        download_name=f"sap_cr_list_{timestamp}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/clients/<client_id>/projects/<project_id>/crs/<cr_id>/status")
def update_cr_status_api(client_id: str, project_id: str, cr_id: str):
    payload = request.get_json(silent=True) or {}
    status = normalize_status(payload.get("status", ""))

    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        return jsonify({"ok": False, "message": "Cliente non trovato."}), 404

    project = find_project(client, project_id)
    if project is None:
        return jsonify({"ok": False, "message": "Progetto non trovato."}), 404

    cr = find_cr(project, cr_id)
    if cr is None:
        return jsonify({"ok": False, "message": "CR non trovata."}), 404

    cr["status"] = status
    cr["updated_at"] = now_iso()
    save_data(data)

    return jsonify(
        {
            "ok": True,
            "status": status,
            "status_label": STATUS_META[status]["label"],
            "updated_at": cr["updated_at"],
        }
    )


@app.post("/clients")
def add_client():
    client_name = sanitize_text(request.form.get("client_name"))
    if not client_name:
        flash("Inserisci il nome del cliente.", "error")
        return redirect(url_for("index"))

    data = load_data()
    if any(client.get("name", "").lower() == client_name.lower() for client in data["clients"]):
        flash("Esiste gia un cliente con questo nome.", "error")
        return redirect(url_for("index"))

    data["clients"].append(
        {
            "id": new_id(),
            "name": client_name,
            "projects": [],
        }
    )
    save_data(data)
    flash("Cliente aggiunto.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/delete")
def delete_client(client_id: str):
    data = load_data()
    before = len(data["clients"])
    data["clients"] = [client for client in data["clients"] if client.get("id") != client_id]
    if len(data["clients"]) == before:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    save_data(data)
    flash("Cliente eliminato.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/projects")
def add_project(client_id: str):
    project_name = sanitize_text(request.form.get("project_name"))
    if not project_name:
        flash("Inserisci il nome del progetto.", "error")
        return redirect(url_for("index"))

    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    if any(project.get("name", "").lower() == project_name.lower() for project in client.get("projects", [])):
        flash("Esiste gia un progetto con questo nome per il cliente selezionato.", "error")
        return redirect(url_for("index"))

    client["projects"].append(
        {
            "id": new_id(),
            "name": project_name,
            "crs": [],
        }
    )
    save_data(data)
    flash("Progetto aggiunto.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/projects/<project_id>/delete")
def delete_project(client_id: str, project_id: str):
    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    before = len(client.get("projects", []))
    client["projects"] = [project for project in client.get("projects", []) if project.get("id") != project_id]
    if len(client["projects"]) == before:
        flash("Progetto non trovato.", "error")
        return redirect(url_for("index"))

    save_data(data)
    flash("Progetto eliminato.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/projects/<project_id>/crs")
def add_cr(client_id: str, project_id: str):
    cr_key = sanitize_text(request.form.get("cr_key"))
    created_by = sanitize_text(request.form.get("created_by"))
    description = sanitize_text(request.form.get("description"))
    notes = sanitize_text(request.form.get("notes"))

    if not cr_key or not created_by or not description:
        flash("Compila ID CR, utente creatore e descrizione.", "error")
        return redirect(url_for("index"))

    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    project = find_project(client, project_id)
    if project is None:
        flash("Progetto non trovato.", "error")
        return redirect(url_for("index"))

    if any(item.get("cr_key", "").lower() == cr_key.lower() for item in project.get("crs", [])):
        flash("Esiste gia una CR con questo ID nel progetto selezionato.", "error")
        return redirect(url_for("index"))

    timestamp = now_iso()
    project["crs"].append(
        {
            "id": new_id(),
            "cr_key": cr_key,
            "created_by": created_by,
            "description": description,
            "notes": notes,
            "status": "development",
            "created_at": timestamp,
            "updated_at": timestamp,
        }
    )
    save_data(data)
    flash("CR aggiunta in stato Sviluppo.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/projects/<project_id>/crs/<cr_id>/update")
def update_cr(client_id: str, project_id: str, cr_id: str):
    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    project = find_project(client, project_id)
    if project is None:
        flash("Progetto non trovato.", "error")
        return redirect(url_for("index"))

    cr = find_cr(project, cr_id)
    if cr is None:
        flash("CR non trovata.", "error")
        return redirect(url_for("index"))

    cr_key = sanitize_text(request.form.get("cr_key"))
    created_by = sanitize_text(request.form.get("created_by"))
    description = sanitize_text(request.form.get("description"))
    notes = sanitize_text(request.form.get("notes"))
    status = normalize_status(request.form.get("status"))

    if not cr_key or not created_by or not description:
        flash("ID CR, utente creatore e descrizione sono obbligatori.", "error")
        return redirect(url_for("index"))

    if any(item.get("id") != cr_id and item.get("cr_key", "").lower() == cr_key.lower() for item in project.get("crs", [])):
        flash("Nel progetto esiste gia un'altra CR con questo ID.", "error")
        return redirect(url_for("index"))

    cr["cr_key"] = cr_key
    cr["created_by"] = created_by
    cr["description"] = description
    cr["notes"] = notes
    cr["status"] = status
    cr["updated_at"] = now_iso()
    save_data(data)
    flash("CR aggiornata.", "success")
    return redirect(url_for("index"))


@app.post("/clients/<client_id>/projects/<project_id>/crs/<cr_id>/delete")
def delete_cr(client_id: str, project_id: str, cr_id: str):
    data = load_data()
    client = find_client(data, client_id)
    if client is None:
        flash("Cliente non trovato.", "error")
        return redirect(url_for("index"))

    project = find_project(client, project_id)
    if project is None:
        flash("Progetto non trovato.", "error")
        return redirect(url_for("index"))

    before = len(project.get("crs", []))
    project["crs"] = [cr for cr in project.get("crs", []) if cr.get("id") != cr_id]
    if len(project["crs"]) == before:
        flash("CR non trovata.", "error")
        return redirect(url_for("index"))

    save_data(data)
    flash("CR eliminata.", "success")
    return redirect(url_for("index"))


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(debug=APP_DEBUG, use_reloader=APP_DEBUG, host="0.0.0.0", port=APP_PORT)