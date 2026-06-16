# Aegis DIM Wishlist Builder

This repo setup converts a current `Destiny 2_ Endgame Analysis.xlsx` workbook into DIM wishlist text files.

## Required repo layout

Upload the files like this:

```text
.github/workflows/build-aegis-from-excel.yml
scripts/build_aegis_wishlist.py
data/endgame.xlsx
```

Rename your Aegis workbook to:

```text
data/endgame.xlsx
```

## GitHub secret

Create a repository secret named:

```text
BUNGIE_API_KEY
```

GitHub path:

```text
Settings → Secrets and variables → Actions → New repository secret
```

## Run it

Go to:

```text
Actions → Build Aegis DIM Wishlist From Excel → Run workflow
```

The workflow will generate:

```text
Aegis-Endgame-S.txt
Aegis-Endgame-A.txt
Aegis-Endgame-A-and-S.txt
Aegis-Unresolved-Names.csv
```

## DIM URL

After the action runs, paste this into DIM:

```text
https://raw.githubusercontent.com/j-courtright7/dim-wishlist/main/Aegis-Endgame-A-and-S.txt
```

## Notes

The generated wishlist uses strict perk-pair matching:

```text
weapon + one Perk 1 recommendation + one Perk 2 recommendation
```

It intentionally does not require barrels, mags, masterworks, or origin traits by default. This makes it much better for vault cleanup than an overly broad old wishlist.
