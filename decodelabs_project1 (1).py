# ============================================================
# DecodeLabs - Python Project 1: To-Do List + Sales Dataset
# Batch: 2026 | Domain: Python Development
# Key Skills: Lists, append(), enumerate(), for loops
# Dataset: Product-Sales-Region.xlsx
# ============================================================

import csv
import os

# ============================================================
# PART 1: THE TO-DO LIST (Core Task from PDF)
# IPO Model: Input → Process → Output
# ============================================================

# STORAGE: The empty list — "The Zero of Lists"
my_tasks = []

def add_task(task_name):
    """INPUT phase — appending to list = INSERT INTO database"""
    task_id = len(my_tasks) + 1
    row = {"id": task_id, "task": task_name, "status": "pending"}
    my_tasks.append(row)   # O(1) amortized — dynamic array on Heap
    print(f"  ✅ Task #{task_id} added: '{task_name}'")

def view_tasks():
    """OUTPUT phase — reading list = SELECT * FROM table"""
    if not my_tasks:
        print("  📋 No tasks yet!")
        return
    print("\n  📋 YOUR TO-DO LIST:")
    print("  " + "-" * 45)
    # enumerate() — professional way (not range(len()))
    for index, row in enumerate(my_tasks, start=1):
        icon = "✅" if row["status"] == "done" else "⬜"
        print(f"  {index}. {icon}  {row['task']}")
    print("  " + "-" * 45)
    pending = sum(1 for r in my_tasks if r["status"] == "pending")
    done    = sum(1 for r in my_tasks if r["status"] == "done")
    print(f"  Total: {len(my_tasks)} | Done: {done} | Pending: {pending}")

def mark_done(task_number):
    """PROCESS phase — modifying list = UPDATE in database"""
    if 1 <= task_number <= len(my_tasks):
        my_tasks[task_number - 1]["status"] = "done"
        print(f"  🎉 Task #{task_number} marked as done!")
    else:
        print(f"  ❌ Task #{task_number} not found.")

def delete_task(task_number):
    """PROCESS phase — removing from list = DELETE FROM table"""
    if 1 <= task_number <= len(my_tasks):
        removed = my_tasks.pop(task_number - 1)
        # Re-assign IDs after deletion (like database auto-increment)
        for i, row in enumerate(my_tasks, start=1):
            row["id"] = i
        print(f"  🗑️  Deleted: '{removed['task']}'")
    else:
        print(f"  ❌ Task #{task_number} not found.")


# ============================================================
# PART 2: SALES DATASET ANALYSIS (Using the same List concepts)
# Reading Product-Sales-Region.xlsx data with Python lists
# This is exactly what the PDF teaches: lists ARE databases!
# ============================================================

def load_sales_data(filepath):
    """
    Load Excel/CSV data into a Python list of dictionaries.
    Each dictionary = one table row (just like the PDF showed!)
        Dictionary  →  Table Row
        Dict 'id'   →  Primary Key
        List        →  Full Database Table
    """
    sales_db = []   # Our in-memory database (list of dicts)

    try:
        # Try reading as Excel using openpyxl (no pandas needed)
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        headers = [str(h) for h in rows[0]]

        for i, row in enumerate(rows[1:], start=1):
            record = {"id": i}
            for header, value in zip(headers, row):
                record[header] = value
            sales_db.append(record)   # list.append() = INSERT INTO

        wb.close()
        print(f"  ✅ Loaded {len(sales_db)} sales records into memory.")
        return sales_db, headers

    except Exception as e:
        print(f"  ⚠️  Could not load file: {e}")
        return [], []


def view_sales_sample(sales_db, n=5):
    """Display first N records — like SELECT TOP 5"""
    print(f"\n  📊 SALES DATA SAMPLE (first {n} records):")
    print("  " + "-" * 70)
    for row in sales_db[:n]:   # list slicing = LIMIT in SQL
        print(f"  #{row['id']:04d} | {str(row.get('Date',''))[:10]} | "
              f"Region: {row.get('Region',''):<8} | "
              f"Product: {row.get('Product',''):<10} | "
              f"Qty: {str(row.get('Quantity','')):<4} | "
              f"Total: ₹{row.get('TotalPrice', 0):>10.2f}")
    print("  " + "-" * 70)


def total_sales_by_region(sales_db):
    """
    GROUP BY Region — using a dictionary (just like the PDF's in-memory DB!)
    Dict key = Region name (like Primary Key / Row Key in Google Bigtable)
    """
    region_totals = {}   # Dictionary → distributed map

    for row in sales_db:   # Iterator Protocol — for task in my_tasks
        region = row.get("Region", "Unknown")
        total  = row.get("TotalPrice", 0) or 0
        if region not in region_totals:
            region_totals[region] = 0
        region_totals[region] += total

    print("\n  🌍 TOTAL SALES BY REGION:")
    print("  " + "-" * 40)
    # Sort by total sales descending
    sorted_regions = sorted(region_totals.items(), key=lambda x: x[1], reverse=True)
    for index, (region, total) in enumerate(sorted_regions, start=1):
        print(f"  {index}. {region:<10}  →  ₹{total:>15,.2f}")
    print("  " + "-" * 40)
    grand_total = sum(region_totals.values())
    print(f"  {'GRAND TOTAL':<14}  →  ₹{grand_total:>15,.2f}")


def top_products(sales_db, top_n=5):
    """Find top selling products — using list + dictionary"""
    product_sales = {}

    for row in sales_db:
        product = row.get("Product", "Unknown")
        total   = row.get("TotalPrice", 0) or 0
        qty     = row.get("Quantity", 0) or 0
        if product not in product_sales:
            product_sales[product] = {"revenue": 0, "units": 0}
        product_sales[product]["revenue"] += total
        product_sales[product]["units"]   += qty

    # Sort and get top N — like ORDER BY revenue DESC LIMIT N
    sorted_products = sorted(
        product_sales.items(),
        key=lambda x: x[1]["revenue"],
        reverse=True
    )[:top_n]

    print(f"\n  🏆 TOP {top_n} PRODUCTS BY REVENUE:")
    print("  " + "-" * 55)
    print(f"  {'Rank':<5} {'Product':<12} {'Revenue':>15} {'Units Sold':>12}")
    print("  " + "-" * 55)
    for rank, (product, data) in enumerate(sorted_products, start=1):
        print(f"  {rank:<5} {product:<12} ₹{data['revenue']:>14,.2f} {data['units']:>12,}")
    print("  " + "-" * 55)


def returned_orders_list(sales_db):
    """
    Filter returned orders — using list comprehension (advanced append pattern)
    This is exactly what list.append() does, but in one line!
    """
    # List comprehension = building a new sub-list (Sharding from the PDF!)
    returned = [row for row in sales_db if row.get("Returned") == 1]

    print(f"\n  📦 RETURNED ORDERS (Total: {len(returned)}):")
    print("  " + "-" * 65)
    if returned:
        for row in returned[:10]:   # Show first 10
            print(f"  OrderID: {row.get('OrderID','')} | "
                  f"Product: {row.get('Product',''):<10} | "
                  f"Region: {row.get('Region',''):<8} | "
                  f"Total: ₹{row.get('TotalPrice', 0):>10.2f}")
        if len(returned) > 10:
            print(f"  ... and {len(returned)-10} more returned orders.")
    print("  " + "-" * 65)
    return_rate = (len(returned) / len(sales_db)) * 100
    print(f"  Return Rate: {return_rate:.2f}%")


def salesperson_performance(sales_db):
    """Rank salespersons by total revenue generated"""
    performance = {}

    for row in sales_db:
        sp    = row.get("Salesperson", "Unknown")
        total = row.get("TotalPrice", 0) or 0
        if sp not in performance:
            performance[sp] = {"revenue": 0, "orders": 0}
        performance[sp]["revenue"] += total
        performance[sp]["orders"]  += 1

    sorted_sp = sorted(performance.items(), key=lambda x: x[1]["revenue"], reverse=True)

    print("\n  👤 SALESPERSON PERFORMANCE:")
    print("  " + "-" * 55)
    print(f"  {'Rank':<5} {'Name':<12} {'Revenue':>15} {'Orders':>8}")
    print("  " + "-" * 55)
    for rank, (name, data) in enumerate(sorted_sp, start=1):
        print(f"  {rank:<5} {name:<12} ₹{data['revenue']:>14,.2f} {data['orders']:>8,}")
    print("  " + "-" * 55)


# ============================================================
# PART 3: COMBINED MAIN ENGINE
# if __name__ == "__main__": — The Gatekeeper (from the PDF)
# ============================================================

def todo_menu():
    """To-Do List interactive menu"""
    print("\n" + "=" * 55)
    print("  📝 PART 1: TO-DO LIST MANAGER")
    print("=" * 55)

    # Pre-load some tasks to demonstrate
    add_task("Study Python Lists and append()")
    add_task("Complete DecodeLabs Project 1")
    add_task("Analyse Product-Sales-Region dataset")
    add_task("Submit project before deadline")

    print()
    view_tasks()

    # Demonstrate mark done and delete
    print("\n  [Marking Task 1 as done...]")
    mark_done(1)
    print("\n  [Marking Task 3 as done...]")
    mark_done(3)

    print()
    view_tasks()


def sales_menu(sales_db):
    """Sales data analysis menu"""
    print("\n" + "=" * 55)
    print("  📊 PART 2: PRODUCT SALES DATASET ANALYSIS")
    print("=" * 55)

    view_sales_sample(sales_db, n=5)
    total_sales_by_region(sales_db)
    top_products(sales_db, top_n=5)
    returned_orders_list(sales_db)
    salesperson_performance(sales_db)


def main():
    print("\n" + "=" * 55)
    print("  🚀 DecodeLabs — Python Project 1")
    print("  Batch 2026 | Junior Python Developer")
    print("=" * 55)

    # ── PART 1: To-Do List ──────────────────────────────────
    todo_menu()

    # ── PART 2: Sales Dataset Analysis ──────────────────────
    # Resolve path to the dataset (same folder or specify path)
    dataset_path = "Product-Sales-Region.xlsx"
    if not os.path.exists(dataset_path):
        # Try common locations
        for candidate in [
            "Product-Sales-Region.xlsx",
            os.path.join(os.path.dirname(__file__), "Product-Sales-Region.xlsx"),
            "/mnt/user-data/uploads/Product-Sales-Region.xlsx",
        ]:
            if os.path.exists(candidate):
                dataset_path = candidate
                break

    print(f"\n  📂 Loading dataset: {os.path.basename(dataset_path)}")
    sales_db, headers = load_sales_data(dataset_path)

    if sales_db:
        sales_menu(sales_db)
    else:
        print("  ⚠️  Place 'Product-Sales-Region.xlsx' in the same folder and re-run.")

    print("\n" + "=" * 55)
    print("  ✅ Project 1 Complete! — DecodeLabs Batch 2026")
    print("  'You aren't just writing a script.")
    print("   You are building the memory for your application.'")
    print("=" * 55 + "\n")


# The Gatekeeper — Project 1 is the Gatekeeper (from PDF)
if __name__ == "__main__":
    main()
