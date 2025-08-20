# Internal Excel Editor (Streamlit)

Bienvenue sur l'application *Internal Excel Editor* !  Cette application
Streamlit permet à vos équipes de modifier des classeurs Excel en ligne via
une interface conviviale, sans installer quoi que ce soit sur leur poste.

## Fonctionnalités

- **Upload de fichier** : sélectionnez un fichier `.xlsx` depuis votre
  ordinateur.  L'application lit toutes les feuilles du classeur.
- **Choix de la feuille** : choisissez la feuille à éditer via une liste
  déroulante.
- **Édition interactive** : modifiez les cellules directement dans un
  tableau éditable.  Si le composant AgGrid est installé, vous pouvez
  également trier et filtrer les colonnes.
- **Téléchargement** : téléchargez le classeur mis à jour en un clic.
  Toutes les autres feuilles non modifiées sont conservées telles quelles.

## Déploiement sur Streamlit Cloud

1. **Forker ou cloner** ce dépôt sur votre compte GitHub.
2. Rendez‑vous sur [Streamlit Cloud](https://share.streamlit.io) et créez
   une nouvelle application en pointant vers ce dépôt et le fichier `app.py`.
3. Lors du premier déploiement, Streamlit installera les dépendances
   listées dans `requirements.txt`.  Patientez quelques minutes.
4. Une fois le déploiement terminé, partagez l'URL générée avec vos
   collaborateurs.

## Exécution en local

Si vous souhaitez tester l'application en local :

```bash
python -m venv .venv
# Sous Windows : .venv\Scripts\activate
# Sous macOS/Linux : source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Puis ouvrez votre navigateur à l'adresse indiquée (par défaut
`http://localhost:8501`).

## Sécurité

Cette application traite des données potentiellement sensibles.  Assurez‑vous
de ne **jamais** publier de classeurs contenant des données confidentielles
dans un dépôt public.  Lorsque vous déployez l'app, restreignez l'accès
à vos utilisateurs authentifiés.