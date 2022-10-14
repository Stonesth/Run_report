from cx_Freeze import setup, Executable


setup(
    name = "run_report",
    version = "0.1",
    description = "",
    executables = [Executable("run_report.py")]
)
