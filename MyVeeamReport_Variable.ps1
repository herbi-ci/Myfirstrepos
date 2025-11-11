# -----------------------------
# MyVeeamReport_Variable.ps1
# Variables et valeurs par défaut pour MyVeeamReport
# Editer ce fichier avant d'exécuter le script principal
# -----------------------------

# Chemin du rapport HTML (sera créé/écrasé)
$ReportPath = "C:\Reports\Veeam_KPI_Report.html"

# Paramètres SMTP (remplir avant d'activer l'envoi)
$SMTPServer = "smtp.yourdomain.local"
$SMTPPort   = 25
$MailFrom   = "veeam-report@yourdomain.local"
$MailTo     = "admin@yourdomain.local"
# Titre du mail (garder comme demandé)
$MailSubject = "Veeam KPI Report – Infrastructure Backup Overview"

# Analyse sur X jours pour calculer les taux (par défaut 30)
$JobsWindowDays = 30

# Envoi du mail : par défaut désactivé (tu peux activer si tu veux)
$SendInBody       = $false
$SendAsAttachment = $false

# Fin du fichier de variables
