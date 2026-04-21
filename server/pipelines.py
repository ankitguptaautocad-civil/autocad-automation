"""Subprocess wrappers that run the existing Python pipeline scripts.

Scripts are executed with cwd=job_dir so their glob-based input discovery
(candidate_search_dirs) finds the uploaded files without any code changes.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


class PipelineError(RuntimeError):
    def __init__(self, message: str, stdout: str = "", stderr: str = ""):
        super().__init__(message)
        self.stdout = stdout
        self.stderr = stderr


def _run(cmd: list[str], cwd: Path) -> subprocess.CompletedProcess:
    result = subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        raise PipelineError(
            f"{cmd[1] if len(cmd) > 1 else cmd[0]} failed with exit code {result.returncode}",
            stdout=result.stdout,
            stderr=result.stderr,
        )
    return result


def run_dxf_pipeline(scripts_dir: Path, job_dir: Path) -> dict:
    """Run dxf_columns_walls_pipeline.py on DXFs already placed in job_dir."""
    script = scripts_dir / "dxf_columns_walls_pipeline.py"
    if not script.exists():
        raise PipelineError(f"Script not found: {script}")

    cmd = [
        sys.executable,
        str(script),
        "--dxf-dir", str(job_dir),
        "--output-dir", str(job_dir),
    ]
    result = _run(cmd, cwd=job_dir)

    outputs = sorted(
        p for p in job_dir.glob("*.xlsx")
        if p.name.endswith(("_col_rectangles_m_v2_wall_assisted.xlsx", "_walls_m_v2.xlsx"))
    )
    return {"stdout": result.stdout, "stderr": result.stderr, "outputs": outputs}


def run_node_pipeline(scripts_dir: Path, job_dir: Path) -> dict:
    """Run Unfiltered then Appended generators in sequence."""
    step1 = scripts_dir / "Unfiltered column coordinates generator.py"
    step2 = scripts_dir / "Appended node coordinate generator.py"
    for s in (step1, step2):
        if not s.exists():
            raise PipelineError(f"Script not found: {s}")

    r1 = _run([sys.executable, str(step1)], cwd=job_dir)
    r2 = _run([sys.executable, str(step2)], cwd=job_dir)

    outputs = sorted(
        p for p in job_dir.glob("*.xlsx")
        if (
            p.name.startswith("node_coordinates_")
            or p.name.startswith("other_coordinates_")
            or p.name.endswith("_column_beam_pairs.xlsx")
        )
    )
    return {
        "stdout": r1.stdout + "\n---\n" + r2.stdout,
        "stderr": r1.stderr + "\n---\n" + r2.stderr,
        "outputs": outputs,
    }
