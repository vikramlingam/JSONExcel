"""Read Excel and create a data-quality profile."""

from jsonexcel import from_excel, profile

profile(from_excel("examples/data/sales.xlsx"), "sales_profile.xlsx")
