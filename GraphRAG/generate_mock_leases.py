import os

os.makedirs("local_cosmos_db/source_files", exist_ok=True)

lease_1 = """
COMMERCIAL OFFICE LEASE AGREEMENT

Lease ID: OFFICE-001
Tenant: Alpha Tech Solutions
Landlord: Metro Properties LLC
Property: Innovation Tower, Floor 12

Term:
The lease shall commence on January 1, 2025 and continue for 5 years.

Rent:
Monthly rent shall be $12,000 payable on the first day of each month.

Late Payment:
Tenant shall receive a 5-day grace period. After the grace period,
a late fee of $250 shall be assessed.

Insurance:
Tenant must maintain general liability insurance coverage of at least
$1,000,000 throughout the lease term.

Maintenance:
Landlord shall maintain common areas.
Tenant shall maintain all equipment located within the leased premises.

Termination:
Either party may terminate the agreement upon material breach if
such breach remains uncured for 30 days after written notice.
"""

lease_2 = """
RETAIL LEASE AGREEMENT

Lease ID: RETAIL-001
Tenant: City Fashion Outlet
Landlord: Urban Retail Holdings
Property: Central Mall Unit 45

Term:
The lease shall commence on March 1, 2025 and continue for 3 years.

Rent:
Monthly rent shall be $8,500.

Late Payment:
Tenant shall receive a 5-day grace period.
A penalty of $300 shall apply thereafter.

Insurance:
Tenant must maintain liability insurance coverage of $1,000,000.

Maintenance:
Tenant shall be responsible for maintaining storefront displays.
Landlord shall maintain mall common areas.

Termination:
Failure to pay rent for 45 consecutive days shall constitute
grounds for lease termination.
"""

lease_3 = """
RESIDENTIAL LEASE AGREEMENT

Lease ID: RES-001
Tenant: Sarah Johnson
Landlord: Green Valley Apartments
Property: Apartment 302

Term:
Lease term shall be 12 months.

Rent:
Monthly rent shall be $1,500.

Late Payment:
A grace period of 3 days shall be allowed.
A late fee of $75 shall be charged thereafter.

Pets:
One domestic cat is permitted.

Maintenance:
Tenant shall maintain cleanliness of the premises.
Landlord shall maintain structural components and utilities.

Termination:
Either party may terminate the lease with 60 days written notice.
"""

leases = {
    "office_lease.txt": lease_1,
    "retail_lease.txt": lease_2,
    "residential_lease.txt": lease_3
}

for filename, content in leases.items():
    path = os.path.join(
        "local_cosmos_db",
        "source_files",
        filename
    )

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

print("Mock leases generated successfully.")