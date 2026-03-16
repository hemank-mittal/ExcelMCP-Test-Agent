"""File Discovery Agent — scans the input directory for Excel workbooks."""

from crewai import Agent
from agents.llm_config import get_llm
from tools.file_scanner import scan_excel_files


def create_file_discovery_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="File Discovery Agent",
        goal=(
            "Scan the specified input directory for Excel workbooks (.xlsx, .xlsm, "
            ".xlsb, .xls) and present a clear list of available files for processing."
        ),
        backstory=(
            "You are a filesystem expert who knows how to locate Excel workbooks "
            "efficiently. You filter out temporary lock files and present results in "
            "a clean, numbered format so the user can easily choose which file to process."
        ),
        tools=[scan_excel_files],
        llm=get_llm(),
        verbose=verbose,
    )
