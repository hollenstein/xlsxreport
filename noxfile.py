import nox


@nox.session(python=["3.9", "3.10", "3.11", "3.12"], tags=["release"])
def installation(session):
    """Test installation of the package."""
    session.install(".[tests]")
    session.run("pytest")


@nox.session(python=["3.9"], tags=["dev"], reuse_venv=True)
def test(session):
    """Run the test suite."""
    session.install(".[tests]")
    session.run("pytest")
