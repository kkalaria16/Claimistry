from app import app, db, Employee, ReimbursementForm, ExpenseEntry
from datetime import datetime, timedelta

def init_db():
    with app.app_context():
        # Drop all tables and recreate them
        db.drop_all()
        db.create_all()
        
        # Add a test employee
        employee = Employee(
            employee_id="EMP001",
            name="John Doe",
            bank_name="Test Bank",
            account_number="1234567890",
            ifsc_code="TEST0001234"
        )
        db.session.add(employee)
        
        # Add a test reimbursement form
        form = ReimbursementForm(
            employee_id="EMP001",
            designation="Software Engineer",
            location="Mumbai",
            from_date=datetime.now() - timedelta(days=30),
            to_date=datetime.now(),
            total_amount=5000.00,
            image_filename="test_form.jpg"
        )
        db.session.add(form)
        db.session.flush()  # To get the form ID
        
        # Add some test expense entries
        expenses = [
            {"date": datetime.now() - timedelta(days=10), "from_location": "Home", "to_location": "Office", "purpose": "Commute", "mode_of_travel": "2-Wheeler", "distance_km": 15.5, "amount_rs": 250.00},
            {"date": datetime.now() - timedelta(days=8), "from_location": "Office", "to_location": "Client Site", "purpose": "Client Meeting", "mode_of_travel": "Cab", "distance_km": 25.0, "amount_rs": 450.00},
            {"date": datetime.now() - timedelta(days=5), "from_location": "Home", "to_location": "Office", "purpose": "Commute", "mode_of_travel": "2-Wheeler", "distance_km": 15.5, "amount_rs": 250.00},
            {"date": datetime.now() - timedelta(days=1), "from_location": "Office", "to_location": "Airport", "purpose": "Business Travel", "mode_of_travel": "Cab", "distance_km": 40.0, "amount_rs": 1200.00},
        ]
        
        for i, exp in enumerate(expenses, 1):
            entry = ExpenseEntry(
                form_id=form.id,
                **exp
            )
            db.session.add(entry)
        
        # Commit all changes
        db.session.commit()
        print("Database initialized with test data!")

if __name__ == "__main__":
    init_db()
