<#
Install pre-commit hooks for the repository and create an initial detect-secrets baseline.
Run this from project root (PowerShell):
  .\tools\install_precommit.ps1
#>
Write-Host "Installing pre-commit and detect-secrets..."
python -m pip install --upgrade pip
python -m pip install pre-commit detect-secrets

Write-Host "Generating detect-secrets baseline (using python -m detect_secrets)..."
python -m detect_secrets > .secrets.baseline

Write-Host "Installing pre-commit hooks (using python -m pre_commit)..."
python -m pre_commit install
Write-Host "Pre-commit hooks installed. Run 'pre-commit run --all-files' to scan the repo." 

Exit 0
