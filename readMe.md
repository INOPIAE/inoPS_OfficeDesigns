# inoPS OfficeDesigns
Eine Sammlung von Scripten zum Arbeiten mit Office Designs und BuilingBlocks.

## Office Designs
### Erstellen von Office Designs

Am besten erstellt man ein Office Design aus PowerPoint heraus, in dem das aktuelle Design einer Präsentation über Entwurf - Designs - Designs - aktuelles Design speichern gespreichert wird

### Speicherort

Die Office Designs sind an 2 Stellen gespeichert.

Benutzerdefinierter Pfad

`C:\Users\USERNAME\AppData\Roaming\Microsoft\Templates\Document Themes`

Systempfad

`C:\Program Files\Microsoft Office\root\Document Themes 16`


### Standardvorlagen für Word, Excel und PowerPoint mit dem CI-Design

1. Erstellen von diesen Basis-Vorlagen:

   PowerPoint-Vorlage: Blank.potx

   Word-Vorlage: Normal.dotm

   Excel-Vorlage: Mappe.xltx

   Der Name der Excel-Vorlage ist abhängig von der Systemsprache unterscheiden kann. Ist die Systemsprache zum Beispiel Englisch ist, den Namen der Excel-Vorlage auf Book.xltx anpassen.

2. Legen Sie nun die PowerPoint- und Word-Vorlage unter dem folgenden Pfad ab:

   `C:\Users\USERNAME\AppData\Roaming\Microsoft\Templates`

3. Legen Sie die Excel-Vorlage unter dem folgenden Pfad ab:

    `C:\Users\USERNAME\AppData\Roaming\Microsoft\Excel\XLSTART`

In Excel und Word muss die Einstellung "Startbildschirm beim Start dieser Anwendung anzeigen" deaktiviert sein. Die Einstellung findet sich unter Datei - Optionen in der Kategorie allgemein.

### Anpassen einer bestehenden Normal.dotx

Die Design-Information wird in der Datei im ZIP-Pfad Datei\word\theme\theme1.xml gespeichert.

Entweder die Datei manuell bearbeiten:
1. die Word-Vorlagendatei in eine ZIP-Datei umwandeln
1. die Datei theme1.xml austauschen 
1. die Word-Vorlagendatei in eine Vorlage umwandeln

Oder die Scriptdatei verwenden.

``` WordNormalDotAnpassen.ps1 -xmlFileOutside ="PFAD\XML-Datei zum Austausch" -$originalFile = "PFAD\Normal.dotm" ```

### Löschen der MS Standard-Designs

Bis auf das Office Theme können alle Themes aus dem Pfad `C:\Program Files\Microsoft Office\root\Document Themes 16` gelöscht werden.

``` OfficeDesignLoeschen.ps1 -keepDesign = "Unternehmen.*" ```

## Building Blocks
### Erstellen von eigenen BuildingBlocks.dotx

Am besten erstellt man ein neue Building Blocks.dotx in dem man die eine bestehende BuildingBlocks.dotx kopiert und umbennent.

### Speicherort

Die Building Blocks Designs sind an 2 Stellen gespeichert.

Benutzerdefinierter Pfad

`C:\Users\USERNAME\AppData\Roaming\Microsoft\Document Building Blocks\1031\16`

Systempfad

`C:\Program Files\Microsoft Office\root\Office16\Document Parts\1031\16`

## Skripte

Die PowerShell-Skripte finden sich im Ordner [scripts](.\scripts).

Die Skripte liegen zweimal vor. Einmal unsigniert und einmal signiert.

Für die signierten Dateien kann es notwendig sein, die Stammzertifikate der benutzten CA in die Truststores von Windows hinzuzufügen.

Die Stammzertifikate befinden sich im Ordner [certificates](.\certificates).

Root zu den `Vertrauenswürdigen Stammzertifizierungsstellen` und die andere Datei zu `Zwischenzertifizierungsstellen`.
