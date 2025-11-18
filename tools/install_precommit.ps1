<#
Install pre-commit hooks for the repository and create an initial detect-secrets baseline.
Run this from project root (PowerShell):
  .\tools\install_precommit.ps1
#>
Write-Host "Installing pre-commit and detect-secrets..."
python -m pip install --upgrade pip
python -m pip install pre-commit detect-secrets

Write-Host "Generating detect-secrets baseline..."
detect-secrets scan > .secrets.baseline

Write-Host "Installing pre-commit hooks..."
pre-commit install
Write-Host "Pre-commit hooks installed. Run 'pre-commit run --all-files' to scan the repo." 

Exit 0
