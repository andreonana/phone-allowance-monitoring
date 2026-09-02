"""
Phone Allowance Monitoring — moteur métier.

Ce package contient toute la logique métier de l'application, indépendante
de l'interface graphique :
    - normalization  : normalisation / validation des numéros de téléphone
    - models         : structures de données (dataclasses)
    - excel_io       : lecture/écriture Excel robustes (détection de colonnes)
    - reference_repository : persistance de la base de référence (collaborateurs, numéros)
    - history_repository    : persistance de l'historique mensuel des consommations
    - reconciliation_engine : rapprochement numéro -> collaborateur
    - consolidation  : agrégation Orange + MTN par collaborateur + comparaison mensuelle
    - alerts_engine  : génération des alertes configurables
    - report_export  : génération du rapport Excel final multi-feuilles
    - config         : paramètres applicatifs (seuils d'alerte, etc.)
    - paths          : résolution des chemins de données (portable Windows/Mac/Linux)

Aucune base de données n'est utilisée : toute la persistance se fait via des
classeurs Excel (.xlsx) stockés dans le dossier `data/` de l'application.
"""
