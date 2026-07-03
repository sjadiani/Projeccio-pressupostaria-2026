# 📊 Seguiment pressupostari d'ingressos — Web per a ens locals

Aplicació web perquè qualsevol ajuntament pugi el seu Excel de seguiment (format ABSIS) i descarregui l'**Excel d'anàlisi validat**, sense instal·lar res ni tocar codi.

## Què genera

Un Excel nou de 4 pestanyes amb **fórmules vives** (canviant el mes o els llindars a PARAMETRES, tot es recalcula sol):

| Pestanya | Contingut |
|---|---|
| `PARAMETRES` | Llindars d'alerta configurables: crític ±40/−60 pp, avís ±20/−40 pp, pendent 5%, **execució alta 95%** (per sobre no salten alertes "per sota"), any i mes analitzat. |
| `ANALISI_INGRESSOS_2026` | Una fila per aplicació: projecció mensual acumulada 2026, execució acumulada (Q), % Real (R), % Esperat pel perfil històric (S), desviació en pp (T), ràtio (U), **semàfor d'anomalia** (V, basat en T), justificació (W) i **metodologia AIC** (AN–AO). |
| `RESUM_ALERTES` | Recompte automàtic per nivell (COUNTIF) i models emprats. |
| `MATRIU_RISCOS` | Nivells d'alerta + tipologia d'anomalies + **detecció concreta** de les 4 anomalies sobre les dades (aplicació, mes i valor). |

## Metodologia (resum)

- **Acumulat net**: les sèries es tracten perquè mai decreixin, neutralitzant devolucions/anul·lacions d'exercicis tancats i regularitzacions de padró.
- **Projecció**: per cada aplicació es proven **ETS (Holt-Winters)** i **SARIMA** i es tria el d'**AIC mínim**; si la sèrie és curta, perfil estacional simple.
- **Perfil esperat**: % acumulat mitjà per mes, **filtrant anys atípics** (2020 i totals anòmals per MAD).
- **Alerta**: semàfor basat en la **desviació en punts (T)** amb els llindars de PARAMETRES, i la regla del **95%** (execució ja pràcticament completa ⇒ sense alertes per sota).

## Fitxers

| Fitxer | Per a què |
|---|---|
| `app.py` | L'aplicació web completa (autònoma). |
| `Colab_Llancar_Web.ipynb` | Per provar-la des de Google Colab en 3 cel·les. |
| `requirements.txt` | Dependències per al desplegament permanent. |

## Posada en marxa

### Opció recomanada per als ajuntaments — Streamlit Community Cloud (URL permanent, gratuïta)
Aquesta és l'opció que evita que els ajuntaments hagin de tocar Colab o codi:
1. Creeu un compte a GitHub i pugeu-hi `app.py` i `requirements.txt` a un repositori.
2. Aneu a [share.streamlit.io](https://share.streamlit.io), inicieu sessió amb GitHub i cliqueu "New app" → seleccioneu el repositori → Deploy.
3. Obtindreu una **URL pública permanent** (p. ex. `https://seguiment-pressupostari.streamlit.app`).
4. Compartiu aquesta URL amb els ajuntaments: només hauran d'**obrir l'enllaç, pujar el seu Excel i descarregar el resultat**.

### Opció de prova — Google Colab
1. Obriu `Colab_Llancar_Web.ipynb` a [Google Colab](https://colab.research.google.com).
2. Executeu les 3 cel·les (al Pas 2, pugeu `app.py`).
3. Obriu l'enllaç que apareix i enganxeu la contrasenya mostrada.

### Opció local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Com la fa servir un ajuntament (un cop desplegada)
1. Obre la URL de la web.
2. Puja el seu Excel de seguiment (.xlsx). L'app detecta sola la fulla d'històric d'ingressos.
3. Clica "Generar l'Excel d'anàlisi" i espera (2–4 minuts per a ~100 aplicacions).
4. Descarrega l'Excel resultant. El fitxer original no es modifica mai.

## Requisits del fitxer d'entrada
La fulla d'històric ha de tenir el format pivot d'ABSIS: una fila amb els blocs "Exercici N-X: YYYY" (un bloc de 10 columnes per any) i, dins de cada bloc, les columnes econòmica, descripció, Mes, previsió inicial, previsió definitiva, DRN i recaptació, amb 12 files de mes per aplicació.
