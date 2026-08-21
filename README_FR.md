# QuickOCR (Version Française)

Un utilitaire OCR portable et hors ligne pour Windows. Capturez et extrayez instantanément du texte à partir d'images, de vidéos, de jeux et d'interfaces non copiables.

[![VirusTotal](https://img.shields.io/badge/VirusTotal-Scan_Result-blue?logo=virustotal)](https://www.virustotal.com/gui/file/b0f5e4cbd0048ef9d6329f7d71e0a05cceda1d62596f63d01b0e3d1489617298/detection)

## Fonctionnalités
* **Capture Visuelle :** Dessinez une zone sur votre écran pour capturer le texte.
* **Bilingue :** Anglais et français inclus, ensemble ou séparément.
* **Ajoutez votre langue :** Déposez un fichier `.traineddata` Tesseract et choisissez-le dans l'application.
* **Filtres Avancés :** Algorithmes spéciaux pour lire le texte sur des fonds bruyants ou colorés (Menus de jeux, Cartes à collectionner).
* **Portable :** Fichier `.exe` unique. Aucune installation. Aucun droit d'administrateur.

## Utilisation
1. Téléchargez `QuickOCR.exe` via le lien ci-dessous.
2. Lancez `QuickOCR.exe` (C'est un fichier unique, sans installation).
3. Cliquez sur **CAPTURE ZONE**.
4. Encadrez le texte à capturer avec votre souris.
5. Le texte est automatiquement copié dans votre presse-papier.

Pour annuler une capture : cliquez une fois sans glisser, faites un clic droit, ou appuyez
sur **Échap**.

## Dépannage
* **L'écran est assombri et rien ne se passe.** C'est la surface de capture. Appuyez sur
  **Échap**, faites un clic droit, ou cliquez une fois pour la fermer.
* **Le texte extrait est incorrect.** La fenêtre de résultat affiche un avertissement
  lorsque Tesseract n'est pas sûr de sa lecture : vérifiez le texte avant de le coller.
* **Le texte dans une autre langue est incorrect** (`Größe` lu `GroBe`, `niño` lu `nino`).
  QuickOCR ne reconnaît que les langues installées. Ajoutez la langue, voir ci-dessous.
* **Signaler un bug.** QuickOCR écrit un journal dans `%APPDATA%\QuickOCR\quickocr.log`.
  Le joindre à un ticket aide grandement au diagnostic.

## Ajouter une langue
1. Téléchargez le fichier `.traineddata` de votre langue depuis
   [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast) (par exemple `deu.traineddata`).
2. Placez-le dans `Tesseract-OCR/tessdata/`.
3. Sélectionnez-le dans le menu **Language** de l'application.

La compilation inclut automatiquement toutes les langues présentes. Tesseract est plus précis
avec une seule langue sélectionnée : préférez une langue unique à une longue combinaison.

## Prérequis
* Windows 10 / 11
* Aucune autre dépendance (Tesseract est inclus).

## Téléchargement
[TÉLÉCHARGER LA DERNIÈRE VERSION](https://github.com/Wolklaw/QuickOCR/releases/latest)

## Licence

**Licence GPLv3**
Ce projet est sous licence **GNU GPLv3**. Voir le fichier [LICENSE](LICENSE) pour les détails.

**En résumé :**
* Vous ne pouvez **pas** fermer le code source et vendre cette application.
* Ce logiciel est gratuit et open source. **Ne payez jamais pour ce logiciel.**
