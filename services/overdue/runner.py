from services.overdue.service import run_overdue_pipeline


def main():
    print("[overdue] pipeline started")
    result = run_overdue_pipeline()
    print("[overdue] pipeline finished")
    items_count = len(result.get("items", []))
    print(f"[overdue] items: {items_count}")


if __name__ == "__main__":
    main()