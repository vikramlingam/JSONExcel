"""Create a data-quality profile workbook."""

from jsonexcel import profile


profile(
    [{"id": 1, "email": "ada@example.com"}, {"id": 2, "email": None}],
    "profile.xlsx",
)
