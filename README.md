# 🤖 Module d'Indexation Automatique RAMEAU pour WinIBW

Ce projet fournit un ensemble de scripts VBScript / JavaScript permettant d'intégrer le service de suggestions de sujet IA du **Hub IdRef (ABES)** directement dans le client de catalogage **WinIBW**.

Il permet aux catalogueurs d'extraire automatiquement les métadonnées d'une notice Unimarc en cours de saisie, d'interroger le moteur d'indexation sujet basé sur l'IA (modèles _Victor_ / LLM), et de récupérer les propositions de vedettes **RAMEAU** soit dans le navigateur web, soit directement injectées dans la notice WinIBW.

---

## 📌 Sommaire

- [Fonctionnalités](#-fonctionnalités)
- [Architecture & Flux de Données](#-architecture--flux-de-données)
- [Fonctions du Script](#-fonctions-du-script)
- [Sécurité & Bonnes Pratiques](#-sécurité--bonnes-pratiques)
- [Installation & Configuration](#-installation--configuration)
- [Utilisation](#-utilisation)
- [Prérequis Techniques](#-prérequis-techniques)

---

## ✨ Fonctionnalités

1. **Extraction automatique Unimarc :**
   - Récupération du PPN (`P3GPP`) et du Login de l'indexeur (`P3GUL`).
   - Extraction des zones de titre : `200 $a` (titre propre), `200 $e` (complément de titre), `200 $i` (titre de partie).
   - Récupération du résumé à partir du texte actuellement sélectionné à l'écran.

2. **Nettoyage automatique des zones RAMEAU :**
   - La fonction `z606withoutDollarsax` assainit les zones `606` existantes en supprimant les sous-zones `$a` et `$x` redondantes lorsqu'un identifiant d'autorité `$3` est présent.

3. **Double mode d'interrogation du Hub IdRef :**
   - **Mode Navigateur (`indexationRameau_html_url`) :** Ouvre la réponse structurée en HTML dans le navigateur par défaut du poste.
   - **Mode Intégré (`indexationRameau_wsh_Secured`) :** Effectue une requête HTTP en arrière-plan et insère directement les propositions sous forme de texte au bas de la notice dans WinIBW.

---

## 🏗 Architecture & Flux de Données

```
┌─────────────────────────┐          ┌──────────────────────────────────┐
│   Notice WinIBW         │          │  API Hub IdRef (ABES)            │
│   (Sudoc / Unimarc)     │          │  https://hub-rdftotal.idref.fr   │
└────────────┬────────────┘          └────────────────┬─────────────────┘
             │                                        │
             │  1. Extraction (PPN, Agent, Titre)     │
             │  2. Sanitize & Encoding UTF-8          │
             ├───────────────────────────────────────►│
             │     Requête HTTP GET (WinHttp)         │
             │                                        │
             │  3. Retour des propositions RAMEAU     │
             │◄───────────────────────────────────────┤
             │                                        │
             │  4. Insertion du résultat              │
             ▼                                        │
   Notice enrichie / IHM
```

---

## 🛠 Fonctions du Script

| Fonction / Procédure             | Description                                                                                                                      |
| :------------------------------- | :------------------------------------------------------------------------------------------------------------------------------- |
| `SanitizeJSString(str)`          | Assainit et échappe les caractères spéciaux, guillemets et retours à la ligne pour éviter toute erreur d'exécution dans JScript. |
| `getTagDollar(tag, dollar)`      | Extrait la valeur d'une sous-zone spécifique (ex: `$a`) d'une zone Unimarc (ex: `200`).                                          |
| `z606withoutDollarsax()`         | Traite la hiérarchie des zones `606` pour supprimer `$a` et `$x` si la zone comporte un `$3`.                                    |
| `indexationRameau_html_url()`    | Prépare l'URL d'appel avec les paramètres encodés et l'ouvre dans le navigateur web par défaut.                                  |
| `indexationRameau_wsh_Secured()` | Effectue un appel synchrone sécurisé via `WinHttp.WinHttpRequest.5.1` et insère le résultat dans la notice.                      |

---

## 🔒 Sécurité & Bonnes Pratiques

Cette version sécurisée du script intègre plusieurs améliorations importantes par rapport aux implémentations VBScript historiques :

- **Transport Chiffré (HTTPS) :** Toutes les communications avec l'API ABES transitent par le protocole sécurisé TLS/HTTPS (`https://hub-rdftotal.idref.fr`).
- **Utilisation de `WinHttp` :** Utilisation de l'API système Windows `WinHttp.WinHttpRequest.5.1` au lieu de `Microsoft.XMLHTTP` (plus stable et plus strict sur la gestion des certificats SSL).
- **Assainissement des données (_Sanitization_) :** Protection contre les injections ou les crashs d'interpréteur dus à la présence de guillemets, d'anti-slashes ou de retours à la ligne dans le texte sélectionné.
- **Encodage UTF-8 :** Forçage du charset UTF-8 lors des requêtes HTTP pour préserver la qualité des caractères accentués et diacritiques.
- **Gestion d'erreurs non intrusive :** Si l'API renvoie un statut d'erreur (ex: 500 ou problème réseau), aucun texte corrompu n'est inséré dans la notice bibliographique. L'utilisateur est informé via une pop-up d'avertissement.

---

## 💻 Installation & Configuration

### 1. Intégration dans WinIBW

1. Ouvrez **WinIBW**.
2. Accédez à la gestion des scripts VBA / VBScript (ex: `Option -> Macros / Scripts`).
3. Copiez-collez le contenu du script dans le fichier de macros utilisateur ou système (ex: `winibw.vbs` ou fichier de fonctions de l'établissement).

### 2. Variables de paramétrage

Le nombre de propositions demandées aux modèles IA peut être ajusté au début du script via la variable `resultsCount` :

```vb
Dim resultsCount
resultsCount = "6" ' Nombre de résultats renvoyés par modèle
```

---

## 🚀 Utilisation

1. En cours de catalogage dans WinIBW, ouvrez la notice à indexer.
2. (Optionnel) **Sélectionnez à la souris** la portion de texte faisant office de résumé dans la notice.
3. Lancez la macro de votre choix :
   - **Par le navigateur :** Exécutez `indexationRameau_html_url`. Une page Web s'ouvrira avec les concepts RAMEAU suggérés.
   - **Directement dans WinIBW :** Exécutez `indexationRameau_wsh_Secured`. Après quelques secondes, le résultat sera ajouté au bas de votre notice :
     ```text
     ... Suggestions :
     - Histoire de France -- 20e siècle
     - Politique et gouvernement -- France -- 1945-1958
     ```

---

## 📋 Prérequis Techniques

- **Poste client :** Windows 10 / 11 avec le logiciel **WinIBW 3.x** installé.
- **Composants Windows :** `MSScriptControl.ScriptControl` et `WinHttp.WinHttpRequest.5.1`.
- **Réseau :** Accès HTTPS vers `hub-rdftotal.idref.fr` (Port 443 autorisé sur le pare-feu).
