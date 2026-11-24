/* Contenu de l'Atelier Excel & Power BI — Rhôna Distribution.
   Pour ajouter une séance : compléter PROGRAMME[n].ready = true et ajouter CONTENT.sXX. */

/* Petite grille façon Excel pour illustrer les notions.
   cols : en-têtes de colonnes (lettres), rows : lignes de valeurs, opts.hi : cellules surlignées ("B2"), opts.fx : barre de formule */
function G(rows, opts) {
  opts = opts || {};
  const hi = opts.hi || [], src = opts.src || [], start = opts.start || 1;
  const n = Math.max(...rows.map(r => r.length));
  const letters = Array.from({length: n}, (_, i) => String.fromCharCode(65 + (opts.col0 || 0) + i));
  let h = '<figure class="xl">';
  if (opts.title) h += '<figcaption>' + opts.title + '</figcaption>';
  if (opts.fx) h += '<div class="fx"><span class="nm">' + (opts.cell || 'fx') + '</span><code>' + opts.fx + '</code></div>';
  h += '<div class="scroll"><table><thead><tr><th></th>' + letters.map(l => '<th>' + l + '</th>').join('') + '</tr></thead><tbody>';
  rows.forEach((r, i) => {
    h += '<tr><th>' + (start + i) + '</th>';
    for (let j = 0; j < n; j++) {
      const ref = letters[j] + (start + i);
      const v = r[j] === undefined ? '' : r[j];
      const cls = (hi.includes(ref) ? 'hi ' : '') + (src.includes(ref) ? 'src ' : '') + (typeof v === 'number' || /^[\d\s.,%€-]+$/.test(String(v)) && v !== '' ? 'num' : '');
      h += '<td class="' + cls + '">' + v + '</td>';
    }
    h += '</tr>';
  });
  return h + '</tbody></table></div></figure>';
}
const F = s => '<pre class="formula"><code>' + s + '</code></pre>';
const TIP = (t, b) => '<aside class="tip"><b>' + t + '</b>' + b + '</aside>';
const TRAP = (t, b) => '<aside class="trap"><b>' + t + '</b>' + b + '</aside>';

const PROGRAMME = [
  {n:1, tool:'Excel', title:"Le CA monte, mais la marge baisse", mission:"Nettoyer l'export des ventes et trouver pourquoi la marge baisse", skills:["Nettoyage","RECHERCHEV","INDEX/EQUIV","SOMME.SI.ENS","Références mixtes"], ready:true},
  {n:2, tool:'Excel', title:"Le comité commercial", mission:"Préparer le comité commercial mensuel", skills:["Tableaux croisés dynamiques","Segments","Remise pondérée"]},
  {n:3, tool:'Excel', title:"Les clients qui ne paient pas", mission:"Balance âgée et liste de relances", skills:["Calculs de dates","Tranches de retard","Filtres avancés"]},
  {n:4, tool:'Excel', title:"Ce qui compte vraiment", mission:"Classer clients et produits, suivre la rotation des stocks", skills:["Analyse ABC","Pareto","Couverture de stock"]},
  {n:5, tool:'Excel', title:"Budget contre réel", mission:"Comparer le budget 2026 au réel et expliquer les écarts", skills:["Écarts volume / prix","Références mixtes"]},
  {n:6, tool:'Excel', title:"Les notes de frais", mission:"Contrôler les notes de frais et repérer les anomalies", skills:["Doublons","Plafonds","Mise en forme conditionnelle"]},
  {n:7, tool:'Excel', title:"À partir de quand on gagne de l'argent", mission:"Compte de résultat, charges fixes / variables, seuil de rentabilité", skills:["SIG","Marge sur coûts variables","Point mort"]},
  {n:8, tool:'Excel', title:"Le tableau de bord du DAF", mission:"Tableau de bord mensuel et simulation d'un plafond de remise", skills:["KPI","Graphiques","Valeur cible","Scénarios"]},
  {n:9, tool:'Power Query', title:"Fini le copier-coller", mission:"Automatiser l'import des exports mensuels", skills:["Power Query","Étapes appliquées"]},
  {n:10, tool:'Power BI', title:"Premiers pas dans Power BI", mission:"Premiers visuels sur les ventes", skills:["Interface","Import","Visuels de base"]},
  {n:11, tool:'Power BI', title:"Relier les tables", mission:"Relier ventes, produits, clients et calendrier", skills:["Modèle de données","Schéma en étoile"]},
  {n:12, tool:'Power BI', title:"Les mesures DAX", mission:"Calculer CA, marge et taux de marge", skills:["SUM","DIVIDE","CALCULATE"]},
  {n:13, tool:'Power BI', title:"Comparer avec l'an dernier", mission:"Cumul annuel et comparaison N-1", skills:["DAX temporel","Table calendrier"]},
  {n:14, tool:'Power BI', title:"Le rapport commercial", mission:"Rapport commercial interactif", skills:["Segments","Filtres","Exploration"]},
  {n:15, tool:'Power BI', title:"La trésorerie", mission:"Prévision de trésorerie et suivi des encaissements", skills:["Rapport de pilotage"]},
  {n:16, tool:'Synthèse', title:"Présenter au DAF", mission:"Présenter le reporting et en faire un projet portfolio", skills:["Mise en page","Publication","Pitch d'entretien"]}
];

const CONTENT = {};

CONTENT.s01 = {
  file: 'files/Rhona_S01.xlsx',
  objectifs: [
    "Nettoyer un export brut (espaces, nombres en texte, doublons)",
    "Aller chercher une information dans une autre table (RECHERCHEV, INDEX + EQUIV)",
    "Calculer CA net, marge et taux de marge ligne par ligne",
    "Construire un tableau de synthèse avec une seule formule (SOMME.SI.ENS + références mixtes)",
    "Transformer des chiffres en message clair pour un DAF"
  ],
  cours: [
  {id:'refs', t:"1. Les références : relatives, absolues, mixtes", d:25, html:
  `<p>Dans une formule, <code>B4</code> ne veut pas dire « la cellule B4 » pour toujours. Excel la comprend comme une <b>position relative</b> : « la cellule située 1 colonne à gauche et 2 lignes au-dessus de moi », par exemple. C'est pour ça que, quand tu recopies une formule, les références bougent toutes seules.</p>
  <h4>La règle de base : ce qui bouge quand on recopie</h4>
  <table class="plain"><thead><tr><th>Tu recopies…</th><th>Ce qui change dans la formule</th><th>Ce qui ne change jamais</th></tr></thead><tbody>
  <tr><td>vers le <b>bas</b> (ou le haut)</td><td>les <b>numéros de ligne</b> (4 → 5 → 6…)</td><td>les lettres de colonne</td></tr>
  <tr><td>vers la <b>droite</b> (ou la gauche)</td><td>les <b>lettres de colonne</b> (B → C → D…)</td><td>les numéros de ligne</td></tr></tbody></table>
  ${TRAP("À retenir absolument","Recopier vers le bas ne fera JAMAIS passer une référence de la colonne K à la colonne L. Si tes valeurs sont rangées à l'horizontale (une par colonne) et que tu recopies à la verticale, aucun $ ne t'aidera : il faut une autre méthode (voir plus bas, cas n°4).")}
  <h4>Le $ : un cadenas</h4>
  <p>Le <code>$</code> se place <b>devant</b> ce que tu veux bloquer : devant la lettre pour bloquer la colonne, devant le chiffre pour bloquer la ligne.</p>
  <table class="plain"><thead><tr><th>Écriture</th><th>Nom</th><th>Recopiée vers le bas</th><th>Recopiée vers la droite</th></tr></thead><tbody>
  <tr><td><code>B1</code></td><td>Relative</td><td>B2, B3, B4…</td><td>C1, D1, E1…</td></tr>
  <tr><td><code>$B$1</code></td><td>Absolue</td><td>$B$1 (fixe)</td><td>$B$1 (fixe)</td></tr>
  <tr><td><code>B$1</code></td><td>Mixte, ligne bloquée</td><td>B$1 (fixe)</td><td>C$1, D$1… (la colonne bouge)</td></tr>
  <tr><td><code>$B1</code></td><td>Mixte, colonne bloquée</td><td>$B2, $B3… (la ligne bouge)</td><td>$B1 (fixe)</td></tr></tbody></table>
  <p>Une astuce pour lire une référence mixte : <b>ce qui est juste après le $ est figé</b>. Dans <code>B$1</code>, le 1 est figé ; dans <code>$B1</code>, le B est figé.</p>

  <h4>Cas n°1 : une seule valeur pour tout le tableau → $B$1</h4>
  <p>Un taux de TVA unique en B1, des prix HT en colonne A. On veut la TVA de chaque ligne. Sans $, la formule <code>=A4*B1</code> recopiée en B5 devient <code>=A5*B2</code> : B2 est vide, donc le résultat vaut 0. Il faut bloquer B1.</p>
  ${G([["Taux TVA","20%"],["",""],["Prix HT","TVA"],["100","20"],["250","50"],["80","16"]],{fx:"=A4*$B$1",cell:"B4",hi:["B4","B5","B6"],src:["B1"]})}
  <p>Recopiée en B5 : <code>=A5*$B$1</code>, en B6 : <code>=A6*$B$1</code>. A4 bouge (chaque ligne a son prix), $B$1 reste fixe (tout le monde utilise le même taux).</p>

  <h4>Cas n°2 : une ligne de paramètres, recopiée vers la droite → B$1</h4>
  <p>Tu as des taux de TVA en ligne 1 (B1, C1, D1) et un seul prix en A3. Tu recopies <b>vers la droite</b> : il faut que la colonne avance (B → C → D) mais que la ligne 1 reste fixe.</p>
  ${G([["","20%","10%","5,5%"],["","","",""],["100","20","10","5,5"]],{fx:"=$A3*B$1",cell:"B3",hi:["B3","C3","D3"],src:["B1","C1","D1","A3"]})}
  <p>En C3, la formule devient <code>=$A3*C$1</code>, en D3 <code>=$A3*D$1</code>. Le prix reste en colonne A grâce à <code>$A</code>, le taux avance de colonne en colonne et reste en ligne 1 grâce à <code>$1</code>.</p>

  <h4>Cas n°3 : un tableau à double entrée → $A3 et B$1 ensemble</h4>
  <p>C'est le cas le plus puissant. Plusieurs prix en colonne A, plusieurs taux en ligne 1, et on veut <b>chaque prix × chaque taux</b>. Une seule formule en B3, recopiée vers la droite <b>puis</b> vers le bas, remplit tout le tableau.</p>
  ${G([["Prix \\ Taux","20%","10%","5,5%"],["","","",""],["100","20","10","5,5"],["250","50","25","13,75"],["80","16","8","4,4"]],{fx:"=$A3*B$1",cell:"B3",hi:["B3","C3","D3","B4","C4","D4","B5","C5","D5"],src:["A3","A4","A5","B1","C1","D1"]})}
  <p>Vérifions la cellule D5 : la formule y devient <code>=$A5*D$1</code>, soit 80 × 5,5 % = 4,4. La colonne A reste fixe pour les prix, la ligne 1 reste fixe pour les taux.</p>
  ${TIP("La question à se poser","Pour chaque référence, demande-toi : « quand je vais recopier vers la droite, est-ce que cette colonne doit bouger ? » et « quand je vais recopier vers le bas, est-ce que cette ligne doit bouger ? ». Si la réponse est non, mets un $ devant.")}

  <h4>Cas n°4 : valeurs à l'horizontale, recopie à la verticale → INDEX</h4>
  <p>C'est le piège de l'exemple ci-dessous : des taux rangés en ligne (K2 à Q2) et des montants rangés en colonne (K5 à K11). On veut que le 1er montant utilise le 1er taux, le 2e montant le 2e taux, etc.</p>
  ${G([["TVA","TVA","TVA"],["20%","18%","16%"],["","",""],["Montant HT","TVA",""],["100","20",""],["150","27",""],["200","32",""]],{col0:10,fx:"=K5*INDEX($K$2:$Q$2;LIGNES($L$5:L5))",cell:"L5",hi:["L5","L6","L7"],src:["K2","L2","M2"]})}
  <p><code>=K5*K$2</code> donnerait 20 % partout, car K ne devient jamais L en descendant. On utilise donc INDEX (voir la notion 4) pour aller chercher le n-ième taux : <code>LIGNES($L$5:L5)</code> vaut 1 en L5, 2 en L6, 3 en L7… car la plage s'agrandit d'une ligne à chaque recopie.</p>
  <p>Solution encore plus simple : <b>réorganiser les données</b>. Mets les taux en colonne, à côté des montants, et écris <code>=K5*M5</code>. En contrôle de gestion, on préfère presque toujours « une ligne = une opération complète ».</p>
  ${TIP("Astuce : la touche F4","Clique sur une référence dans la barre de formule et appuie sur F4 : B1 → $B$1 → B$1 → $B1 → B1. Pas besoin de taper les $ à la main.")}
  ${TRAP("Comment trouver une erreur de $","Si un tableau affiche des zéros ou des résultats bizarres après recopie, clique sur une cellule en bas à droite du tableau et appuie sur F2 : Excel encadre en couleur les cellules utilisées. Tu vois tout de suite quelle référence a glissé au mauvais endroit.")}
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice 1 · Commission des commerciaux</summary><p>Le taux de commission (3 %) est en E1. Les ventes de chaque commercial sont en B2:B10. Quelle formule écrire en C2 pour la recopier jusqu'en C10 ?</p><details><summary>Voir la solution</summary><p><code>=B2*$E$1</code> : la vente change à chaque ligne, le taux est unique donc bloqué en ligne et en colonne.</p></details></details>
  <details class="exo"><summary>Exercice 2 · Prix après remise</summary><p>Des prix catalogue sont en A3:A8, des taux de remise en B1:E1 (5 %, 10 %, 15 %, 20 %). Tu veux le prix remisé pour chaque prix et chaque remise. Quelle formule en B3 ?</p><details><summary>Voir la solution</summary><p><code>=$A3*(1-B$1)</code> : le prix est toujours en colonne A ($A), la remise toujours en ligne 1 ($1). On recopie à droite puis en bas.</p></details></details>
  <details class="exo"><summary>Exercice 3 · Part de chaque mois</summary><p>Le CA des 12 mois est en B2:B13 et le total annuel en B14. En C2, tu veux la part du mois dans le total, puis recopier jusqu'en C13. Quelle formule ?</p><details><summary>Voir la solution</summary><p><code>=B2/$B$14</code> (ou <code>=B2/B$14</code>, puisqu'on recopie seulement vers le bas). Sans $, la formule de C3 diviserait par B15, qui est vide : #DIV/0!.</p></details></details></details>`},

  {id:'clean', t:"2. Nettoyer un export avant de l'analyser", d:15, html:
  `<p>En entreprise, les données sortent d'un logiciel (facturation, comptabilité, ERP) et ne sont <b>jamais</b> parfaites. Trois défauts reviennent tout le temps. Ils sont dangereux parce qu'ils ne se voient pas à l'œil nu et qu'Excel ne prévient pas.</p>
  <h4>Défaut n°1 : les espaces en trop</h4>
  <p>« CLI007 » et « CLI007 » suivi d'un espace se ressemblent à l'écran, mais ce sont deux textes différents pour Excel. Toute recherche (RECHERCHEV, SOMME.SI.ENS…) les traitera comme deux clients distincts.</p>
  ${G([["Code brut","NBCAR","Propre"],["CLI007 ","7","CLI007"],["CLI012","6","CLI012"],["  CLI003","8","CLI003"]],{fx:"=SUPPRESPACE(A2)",cell:"C2",hi:["C2","C3","C4"],src:["A2","A4"]})}
  <ul><li><code>=NBCAR(A2)</code> compte les caractères. Un code client fait 6 caractères : tout ce qui dépasse trahit des espaces.</li>
  <li><code>=SUPPRESPACE(A2)</code> supprime les espaces au début et à la fin, et réduit les espaces doubles au milieu à un seul.</li></ul>
  <h4>Défaut n°2 : les nombres stockés en texte</h4>
  <p>Un nombre importé peut arriver « en texte » : il est aligné à gauche, un petit triangle vert apparaît dans le coin de la cellule. SOMME l'ignore sans rien dire, et ton total est faux.</p>
  ${G([["Quantité","ESTTEXTE","Nombre"],["12","FAUX","12"],["8","VRAI","8"],["5","FAUX","5"]],{fx:"=CNUM(A3)",cell:"C3",hi:["B3"],src:["A3"]})}
  <ul><li><code>=ESTTEXTE(A2)</code> renvoie VRAI si la cellule contient du texte.</li>
  <li><code>=CNUM(A2)</code> convertit un texte qui ressemble à un nombre en vrai nombre (sans effet sur un nombre déjà correct).</li>
  <li>Alternative rapide sans formule : sélectionne la colonne, clique sur le losange jaune ⚠ et choisis « Convertir en nombre ».</li></ul>
  <h4>Défaut n°3 : les lignes en double</h4>
  <p>Une facture exportée deux fois gonfle le CA. Pour les repérer : Accueil &gt; Mise en forme conditionnelle &gt; Règles de mise en surbrillance &gt; Valeurs en double. Pour les supprimer : Données &gt; Supprimer les doublons, <b>en cochant toutes les colonnes</b> (sinon Excel supprime des lignes qui partagent seulement un numéro de facture, alors qu'une facture peut avoir plusieurs produits).</p>
  ${TRAP("Toujours compter avant et après","Note le nombre de lignes avant de nettoyer (=NBVAL(A:A)-1) puis après. Tu sais exactement ce que tu as supprimé, et tu peux le dire au DAF.")}
  ${TIP("Bonne pratique","Ne modifie pas les colonnes d'origine : ajoute une colonne « propre » à côté, avec la formule. Tu gardes une trace de ce qui a été corrigé, et un contrôleur peut toujours vérifier.")}
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice · Le total qui ne tombe pas juste</summary><p>Ta colonne de quantités affiche visuellement 10 + 20 + 30, mais SOMME renvoie 40. Que se passe-t-il et comment le vérifier ?</p><details><summary>Voir la solution</summary><p>Une des valeurs (ici 20) est stockée en texte : SOMME l'ignore. <code>=ESTTEXTE()</code> sur chaque cellule le confirme. On corrige avec <code>=CNUM()</code> ou « Convertir en nombre ».</p></details></details>`},

  {id:'rechv', t:"3. RECHERCHEV : aller chercher une info dans une autre table", d:20, html:
  `<p>Ton export de ventes contient seulement la référence produit (<code>PEM-002</code>). Le prix et la famille sont dans une autre table, l'onglet Produits. RECHERCHEV fait le lien entre les deux, comme quand tu cherches un nom dans un annuaire pour trouver son numéro.</p>
  ${F("=RECHERCHEV(valeur_cherchée ; table ; n°_colonne ; FAUX)")}
  <table class="plain"><thead><tr><th>Argument</th><th>Ce que c'est</th><th>Dans l'exemple</th></tr></thead><tbody>
  <tr><td>valeur_cherchée</td><td>Ce que tu connais déjà</td><td>D2, la référence produit de la vente</td></tr>
  <tr><td>table</td><td>Où chercher. La valeur cherchée doit être dans la <b>1re colonne</b> de cette plage.</td><td><code>Produits!$A:$C</code></td></tr>
  <tr><td>n°_colonne</td><td>Quelle colonne renvoyer, en comptant depuis la 1re colonne de la plage (1, 2, 3…)</td><td>3, la famille</td></tr>
  <tr><td>FAUX</td><td>Correspondance exacte</td><td>Toujours FAUX (ou 0) pour un code</td></tr></tbody></table>
  ${G([["Réf produit","Désignation","Famille","Prix d'achat","Prix de vente"],["VAI-001","Assiette plate (lot de 6)","Vaisselle","14,20","23,90"],["PEM-002","Blender pro 2 L","Petit électroménager","215,00","299,00"],["EMB-001","Boîte repas kraft","Emballages","38,50","59,00"]],{title:"Onglet Produits : colonne 1 = A, colonne 2 = B, colonne 3 = C…",src:["A3"],hi:["C3"]})}
  ${F("=RECHERCHEV(D2;Produits!$A:$C;3;FAUX)   →   Petit électroménager\n=RECHERCHEV(D2;Produits!$A:$E;5;FAUX)   →   299,00")}
  <p>Étape par étape : Excel prend la valeur de D2 (PEM-002), descend la colonne A de Produits jusqu'à la trouver (ligne 3), puis se décale de 3 colonnes vers la droite en comptant A comme la 1re : il renvoie C3.</p>
  <p>Pourquoi les <code>$</code> sur <code>Produits!$A:$C</code> ? Parce qu'on recopie la formule vers le bas sur 1 400 lignes : la table de recherche, elle, ne doit pas bouger (notion 1, cas n°1).</p>
  ${TRAP("Le piège n°1 en entreprise : oublier FAUX","Sans FAUX, Excel fait une recherche « approchée » : il renvoie une valeur proche, pas forcément la bonne, sans aucune erreur. Des reportings entiers sont faux à cause de ça.")}
  <h4>Lire les erreurs</h4>
  <table class="plain"><tbody><tr><td><code>#N/A</code></td><td>Valeur introuvable. Causes possibles : code inexistant, espace parasite, nombre en texte d'un côté et en nombre de l'autre, ou valeur cherchée qui n'est pas dans la 1re colonne de la table.</td></tr>
  <tr><td><code>#REF!</code></td><td>Le n° de colonne dépasse la largeur de la table (ex. 5 avec une plage A:C).</td></tr>
  <tr><td><code>#VALEUR!</code></td><td>Le n° de colonne vaut 0 ou est négatif.</td></tr></tbody></table>
  <h4>SIERREUR : remplacer l'erreur par un message utile</h4>
  ${F('=SIERREUR(RECHERCHEV(D2;Produits!$A:$C;3;FAUX);"Réf inconnue")')}
  <p>SIERREUR teste la formule : si elle renvoie une erreur, il affiche le message à la place. Ici, « Réf inconnue » t'aide à repérer les références absentes du catalogue.</p>
  ${TIP("Astuce","Utilise SIERREUR après avoir compris tes erreurs, pas avant : s'il cache tout dès le début, tu ne vois plus les vrais problèmes de données.")}
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice · Le commercial du client</summary><p>L'onglet Clients contient en A le code client, en B le nom, en C le type, en D la ville, en E le commercial. Ton code client propre est en H2. Quelle formule donne le commercial ?</p><details><summary>Voir la solution</summary><p><code>=RECHERCHEV(H2;Clients!$A:$E;5;FAUX)</code> : la plage commence en A (où sont les codes), et le commercial est dans la 5e colonne de la plage.</p></details></details>`},

  {id:'index', t:"4. INDEX + EQUIV : la version professionnelle", d:20, html:
  `<p>RECHERCHEV a deux limites. La valeur cherchée doit être dans la première colonne, et le n° de colonne est écrit « en dur » : si quelqu'un insère une colonne dans la table, ta formule renvoie la mauvaise donnée sans prévenir. INDEX + EQUIV n'a pas ces défauts. On les apprend séparément, puis ensemble.</p>
  <h4>EQUIV : « à quelle position se trouve… ? »</h4>
  ${F("=EQUIV(valeur ; plage_d'une_seule_colonne_ou_ligne ; 0)")}
  ${G([["Réf"],["VAI-001"],["PEM-002"],["EMB-001"]],{title:"=EQUIV(\"PEM-002\";A1:A4;0) renvoie 3 : c'est la 3e cellule de la plage",hi:["A3"]})}
  <p>Le 0 final demande une correspondance exacte (comme FAUX dans RECHERCHEV).</p>
  <h4>INDEX : « donne-moi la valeur à la position… »</h4>
  ${F("=INDEX(plage ; position)")}
  <p><code>=INDEX(E1:E4;3)</code> renvoie la 3e valeur de la colonne E, soit 299,00. INDEX marche aussi sur une ligne : <code>=INDEX(K2:Q2;2)</code> renvoie la 2e valeur de la ligne, soit 18 % dans l'exemple de la notion 1.</p>
  <h4>Les deux ensemble</h4>
  ${F("=INDEX(Produits!$E:$E;EQUIV(D2;Produits!$A:$A;0))   →   299,00")}
  <p>On lit de l'intérieur vers l'extérieur : EQUIV trouve la position de D2 dans la colonne A (3), puis INDEX prend la 3e valeur de la colonne E.</p>
  ${G([["Réf produit","…","…","Prix d'achat","Prix de vente"],["VAI-001","","","14,20","23,90"],["PEM-002","","","215,00","299,00"]],{title:"EQUIV trouve PEM-002 en position 3 de la colonne A, INDEX renvoie la position 3 de la colonne E",src:["A3"],hi:["E3"]})}
  <table class="plain"><thead><tr><th></th><th>RECHERCHEV</th><th>INDEX + EQUIV</th></tr></thead><tbody>
  <tr><td>Colonne de recherche</td><td>Forcément la 1re</td><td>N'importe laquelle, même à droite</td></tr>
  <tr><td>Insertion d'une colonne</td><td>Casse la formule</td><td>Aucun problème</td></tr>
  <tr><td>Recherche dans une ligne</td><td>Non (il faut RECHERCHEH)</td><td>Oui</td></tr>
  <tr><td>Facilité</td><td>Plus simple à écrire</td><td>Un peu plus long</td></tr></tbody></table>
  ${TIP("Pour l'entretien","Depuis Excel 365, RECHERCHEX remplace les deux : =RECHERCHEX(D2;Produits!A:A;Produits!E:E). Ton Excel 2013 ne l'a pas, mais savoir qu'elle existe est un plus. INDEX + EQUIV marche partout : c'est pour ça qu'on l'apprend d'abord.")}
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice · Chercher vers la gauche</summary><p>Dans une table, les noms des clients sont en colonne B et leurs codes en colonne A. Tu connais le nom (en G2) et tu veux le code. Pourquoi RECHERCHEV ne marche pas, et quelle formule utiliser ?</p><details><summary>Voir la solution</summary><p>RECHERCHEV ne sait chercher que dans la 1re colonne et renvoyer vers la droite. Avec INDEX + EQUIV : <code>=INDEX($A:$A;EQUIV(G2;$B:$B;0))</code>.</p></details></details>`},

  {id:'si', t:"5. SI, ET, OU : faire réagir le tableau", d:15, html:
  `<p>SI pose une question à Excel et affiche une chose ou une autre selon la réponse.</p>
  ${F('=SI(test ; valeur_si_vrai ; valeur_si_faux)')}
  <p>Le test est une comparaison : <code>=</code>, <code>&lt;&gt;</code> (différent), <code>&gt;</code>, <code>&lt;</code>, <code>&gt;=</code>, <code>&lt;=</code>. Les textes s'écrivent entre guillemets, les nombres sans.</p>
  ${F('=SI(C2<20%;"Marge faible";"OK")')}
  <h4>Plusieurs conditions : ET et OU</h4>
  <ul><li><code>ET(test1;test2)</code> est VRAI seulement si <b>tous</b> les tests sont vrais.</li>
  <li><code>OU(test1;test2)</code> est VRAI si <b>au moins un</b> test est vrai.</li></ul>
  ${G([["CA net","Marge","Taux","Alerte"],["1 690","169","10,0%","Marge faible"],["299","84","28,1%",""],["0","0","",""]],{fx:'=SI(ET(A2>0;C2<20%);"Marge faible";"")',cell:"D2",hi:["D2"]})}
  <p>Pourquoi le <code>A2&gt;0</code> ? Pour ne pas signaler les lignes sans chiffre d'affaires (références inconnues), qui n'ont pas de taux de marge.</p>
  <h4>SI imbriqués : plus de deux résultats</h4>
  ${F('=SI(C2<10%;"Critique";SI(C2<20%;"Faible";"Correct"))')}
  <p>Excel teste dans l'ordre : moins de 10 % → Critique ; sinon moins de 20 % → Faible ; sinon Correct. Commence toujours par le cas le plus strict.</p>
  <h4>Éviter la division par zéro</h4>
  ${F('=SI(A2=0;"";B2/A2)')}
  <p>Sans ce SI, une ligne à CA nul afficherait #DIV/0!, et cette erreur se propagerait dans tous tes totaux. <code>""</code> (deux guillemets) affiche une cellule vide.</p>
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice · Prime des commerciaux</summary><p>Un commercial touche une prime si son CA (B2) dépasse 50 000 € <b>et</b> son taux de marge (C2) est d'au moins 30 %. La prime vaut 500 €, sinon 0. Quelle formule ?</p><details><summary>Voir la solution</summary><p><code>=SI(ET(B2>50000;C2>=30%);500;0)</code></p></details></details>`},

  {id:'sommesi', t:"6. SOMME.SI.ENS, NB.SI.ENS, MOYENNE.SI.ENS", d:25, html:
  `<p>Ce sont les fonctions les plus utilisées en contrôle de gestion : additionner, compter ou faire une moyenne <b>seulement pour les lignes qui respectent des critères</b>. C'est ce qui permet de passer de 1 400 lignes de ventes à un tableau de synthèse lisible.</p>
  ${F("=SOMME.SI.ENS(plage_à_additionner ; plage_critère1 ; critère1 ; plage_critère2 ; critère2 ; …)")}
  <p>Excel parcourt les lignes une par une. Pour chaque ligne qui respecte <b>tous</b> les critères, il ajoute la valeur de la plage à additionner.</p>
  ${G([["Commercial","Famille","Mois","CA net"],["Lucas Moreau","Petit électroménager","4","1 352"],["Karim Benali","Vaisselle","4","215"],["Lucas Moreau","Petit électroménager","5","598"],["Lucas Moreau","Emballages","5","472"]],{title:"Combien Lucas a-t-il vendu en petit électroménager ? Seules les lignes 2 et 4 respectent les deux critères.",hi:["D2","D4"]})}
  ${F('=SOMME.SI.ENS(D:D;A:A;"Lucas Moreau";B:B;"Petit électroménager")   →   1 352 + 598 = 1 950')}
  <p>Les plages doivent avoir <b>la même taille</b> (ici des colonnes entières, c'est le plus sûr).</p>
  <table class="plain"><thead><tr><th>Fonction</th><th>Ce qu'elle fait</th><th>Ordre des arguments</th></tr></thead><tbody>
  <tr><td><code>SOMME.SI.ENS</code></td><td>Additionne</td><td>Plage à additionner <b>en premier</b>, puis les couples plage/critère</td></tr>
  <tr><td><code>NB.SI.ENS</code></td><td>Compte les lignes</td><td>Seulement des couples plage/critère</td></tr>
  <tr><td><code>MOYENNE.SI.ENS</code></td><td>Fait la moyenne</td><td>Plage à moyenner en premier, comme SOMME.SI.ENS</td></tr></tbody></table>
  <h4>Écrire les critères</h4>
  <table class="plain"><tbody>
  <tr><td>Texte exact</td><td><code>"Vaisselle"</code>, ou mieux une cellule : <code>$A5</code></td></tr>
  <tr><td>Nombre exact</td><td><code>4</code> ou une cellule : <code>B$4</code></td></tr>
  <tr><td>Comparaison</td><td><code>"&gt;1000"</code>, <code>"&lt;&gt;0"</code> (différent de 0)</td></tr>
  <tr><td>Comparaison avec une cellule</td><td><code>"&gt;="&amp;B1</code> (le &amp; colle le signe et la valeur)</td></tr>
  <tr><td>Commence par</td><td><code>"PEM*"</code> (l'étoile remplace n'importe quelle suite de caractères)</td></tr></tbody></table>
  <h4>La méthode du contrôleur : un tableau entier avec une seule formule</h4>
  <p>On combine SOMME.SI.ENS et les références mixtes (notion 1, cas n°3). Familles en colonne A, numéros de mois en ligne 4 :</p>
  ${G([["Famille \\ Mois","1","2","3"],["Vaisselle","11 795","7 950","12 323"],["Emballages","24 336","26 217","29 524"]],{start:4,fx:"=SOMME.SI.ENS(Ventes!$P:$P;Ventes!$J:$J;$A5;Ventes!$G:$G;B$4)",cell:"B5",hi:["B5","C5","D5","B6","C6","D6"],src:["A5","A6","B4","C4","D4"]})}
  <ul><li><code>Ventes!$P:$P</code>, <code>$J:$J</code>, <code>$G:$G</code> : colonnes de données, totalement bloquées.</li>
  <li><code>$A5</code> : la famille, toujours en colonne A, mais une par ligne.</li>
  <li><code>B$4</code> : le mois, toujours en ligne 4, mais un par colonne.</li></ul>
  ${TRAP("Toujours une ligne de contrôle","Le total de ton tableau de synthèse doit être égal à SOMME de la colonne source. Si une famille est mal orthographiée en colonne A, ses ventes disparaissent du tableau sans erreur : seul le contrôle te le dira.")}
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice 1 · Compter les grosses ventes</summary><p>Combien de lignes de ventes dépassent 1 000 € de CA net (colonne P) ?</p><details><summary>Voir la solution</summary><p><code>=NB.SI.ENS(Ventes!P:P;"&gt;1000")</code></p></details></details>
  <details class="exo"><summary>Exercice 2 · CA d'un client sur un trimestre</summary><p>Le code client est en A2, et tu veux son CA net du 2e trimestre (mois 4 à 6, colonne G).</p><details><summary>Voir la solution</summary><p><code>=SOMME.SI.ENS(Ventes!$P:$P;Ventes!$H:$H;A2;Ventes!$G:$G;"&gt;=4";Ventes!$G:$G;"&lt;=6")</code> : on peut mettre deux critères sur la même colonne pour faire une fourchette.</p></details></details>`},

  {id:'top', t:"7. Faire un classement : GRANDE.VALEUR", d:10, html:
  `${F("=GRANDE.VALEUR(plage ; k)    → la k-ième plus grande valeur\n=PETITE.VALEUR(plage ; k)    → la k-ième plus petite")}
  <p><code>=GRANDE.VALEUR(F2:F31;1)</code> donne le plus gros CA, <code>;2</code> le 2e, etc. Pour un top 5, on écrit les rangs 1 à 5 dans une colonne et on s'en sert comme k.</p>
  <p>GRANDE.VALEUR donne un montant, pas un nom. Pour retrouver le client qui correspond, on réutilise INDEX + EQUIV (notion 4) : EQUIV cherche ce montant dans la colonne des CA, INDEX renvoie le nom à la même position.</p>
  ${G([["Rang","Client","CA net"],["1","Restaurant L'Écullois","32 473"],["2","Restaurant Le Guillotière","26 837"]],{fx:"=INDEX(Clients!$B:$B;EQUIV(C2;Clients!$F:$F;0))",cell:"B2",hi:["B2"],src:["A2","C2"]})}
  <p>En C2 : <code>=GRANDE.VALEUR(Clients!$F:$F;A2)</code>. En B2 : la formule ci-dessus.</p>
  ${TRAP("Limite à connaître","Si deux clients ont exactement le même CA, EQUIV renverra deux fois le premier. Avec des montants au centime près, c'est très rare, mais il faut le savoir.")}`},

  {id:'gestion', t:"8. Les notions de gestion derrière les formules", d:20, html:
  `<p>Excel n'est qu'un outil. Ce qu'on attend de toi en stage, c'est de comprendre ce que tu calcules et de savoir l'expliquer.</p>
  <table class="plain"><tbody>
  <tr><td><b>CA brut</b></td><td>Quantité × prix de vente catalogue</td></tr>
  <tr><td><b>Remise</b></td><td>Réduction accordée au client, en % ou en €</td></tr>
  <tr><td><b>CA net</b></td><td>CA brut − remises : ce que l'entreprise facture vraiment</td></tr>
  <tr><td><b>Coût d'achat</b></td><td>Quantité × prix d'achat fournisseur</td></tr>
  <tr><td><b>Marge commerciale</b></td><td>CA net − coût d'achat : ce qui reste pour payer les salaires, le loyer… et faire du bénéfice</td></tr></tbody></table>
  <h4>Taux de marge ou taux de marque ? La question qui tombe en entretien</h4>
  ${G([["Prix d'achat","Prix de vente","Marge","Taux de marge","Taux de marque"],["100","125","25","25 %","20 %"]],{hi:["D2","E2"]})}
  <ul><li><b>Taux de marge</b> = marge / prix d'achat = 25 / 100 = 25 %</li>
  <li><b>Taux de marque</b> = marge / prix de vente = 25 / 125 = 20 %</li></ul>
  <p>Dans beaucoup d'entreprises, on dit « taux de marge » pour marge / CA (qui est en réalité le taux de marque). Dans ce cas pratique, on suit cet usage. Mais montrer que tu connais la nuance fait bonne impression.</p>
  <h4>Pourquoi une remise coûte si cher</h4>
  <p>Un produit acheté 215 € et vendu 299 € rapporte 84 €. Avec 20 % de remise, il est vendu 239 € : il ne rapporte plus que 24 €. La remise ne retire pas 20 % de la marge, elle en retire <b>71 %</b>, car elle s'applique au prix de vente entier alors que le coût d'achat, lui, ne bouge pas.</p>
  ${G([["","Sans remise","Remise 20 %"],["Prix de vente","299","239"],["Prix d'achat","215","215"],["Marge","84","24"]],{hi:["C4"]})}
  <h4>L'effet mix</h4>
  <p>Le taux de marge global est une moyenne pondérée des taux de chaque famille. Si une famille à bonne marge prend plus de poids, le taux global monte, même si une autre famille s'effondre. C'est pour ça qu'un contrôleur <b>découpe toujours</b> : par famille, puis par commercial, puis par mois, jusqu'à trouver la cause.</p>
  <h4>À toi de jouer</h4>
  <details class="exo"><summary>Exercice · Calcul de marge</summary><p>Un produit est acheté 40 € et vendu 60 €. Calcule le taux de marge et le taux de marque. Puis : quel est le taux de marque si on accorde 10 % de remise ?</p><details><summary>Voir la solution</summary><p>Marge 20 € : taux de marge 50 %, taux de marque 33,3 %. Avec 10 % de remise : prix 54 €, marge 14 €, taux de marque 25,9 %.</p></details></details>`},

  {id:'raccourcis', t:"9. Les raccourcis qui font gagner du temps", d:5, html:
  `<p>En entreprise, la vitesse compte. Ces raccourcis sont ceux qu'on voit le plus chez les contrôleurs de gestion.</p>
  <table class="plain"><tbody>
  <tr><td><kbd>Ctrl</kbd> + <kbd>↓</kbd></td><td>Aller à la dernière ligne remplie</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>Maj</kbd> + <kbd>↓</kbd></td><td>Sélectionner jusqu'à la dernière ligne</td></tr>
  <tr><td>Double-clic sur la poignée</td><td>Recopier une formule jusqu'en bas du tableau</td></tr>
  <tr><td><kbd>F4</kbd></td><td>Ajouter ou changer les $ d'une référence</td></tr>
  <tr><td><kbd>F2</kbd></td><td>Modifier la cellule et voir ses références en couleur</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>Maj</kbd> + <kbd>L</kbd></td><td>Activer ou retirer les filtres</td></tr>
  <tr><td><kbd>Alt</kbd> + <kbd>Entrée</kbd></td><td>Retour à la ligne dans une cellule</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>;</kbd></td><td>Insérer la date du jour</td></tr></tbody></table>
  ${TIP("Astuce","Affichage > Figer les volets > Figer la ligne supérieure : les titres restent visibles quand tu descends dans 1 400 lignes.")}`}
  ],

  quiz: [
    {q:"Tu recopies =A4*B1 vers le bas et tu obtiens des zéros. Que faut-il écrire ?", o:["=A4*B1","=$A$4*B1","=A4*$B$1","=A$4*B1"], a:2, e:"Le taux est dans une seule cellule : il faut bloquer B1 en ligne et en colonne, donc $B$1."},
    {q:"Dans un tableau familles (colonne A) × mois (ligne 4), quelle référence pour la famille ?", o:["A5","$A5","A$5","$A$5"], a:1, e:"La famille est toujours en colonne A (on bloque A) mais change à chaque ligne (on laisse la ligne libre) : $A5."},
    {q:"RECHERCHEV renvoie #N/A alors que le code client existe bien dans la table. Cause la plus probable ?", o:["Le 4e argument est FAUX","Un espace parasite dans le code","La table est trop grande","Il manque un $"], a:1, e:"« CLI007 » et « CLI007 » suivi d'un espace sont deux textes différents. SUPPRESPACE règle le problème."},
    {q:"Dans SOMME.SI.ENS, quelle plage vient en premier ?", o:["La plage du premier critère","La plage à additionner","Le critère","Peu importe"], a:1, e:"SOMME.SI.ENS et MOYENNE.SI.ENS commencent par la plage à additionner (ou à moyenner), puis les couples plage / critère."},
    {q:"Achat 100 €, vente 125 €. Quel est le taux de marque ?", o:["25 %","20 %","125 %","80 %"], a:1, e:"Taux de marque = marge / prix de vente = 25 / 125 = 20 %. Le taux de marge (marge / prix d'achat) serait de 25 %."},
    {q:"Pourquoi préférer INDEX + EQUIV à RECHERCHEV ?", o:["C'est plus court","Ça marche même si on insère une colonne","C'est plus rapide à taper","RECHERCHEV n'existe plus"], a:1, e:"INDEX + EQUIV désigne directement la colonne à renvoyer : insérer une colonne ne la décale pas, et la colonne de recherche peut être n'importe où."}
  ],
  cas: {
    brief: "C'est ton premier jour chez Rhôna Distribution. M. Faure, le DAF, te transmet l'export brut des ventes de janvier à juin 2026.",
    quote: "« Nos ventes ont bien progressé depuis le printemps, mais j'ai l'impression qu'on gagne moins sur chaque vente. Peux-tu remettre ce fichier au propre et me dire d'où vient le problème ? »",
    etapes: [
      {k:'A', t:"Nettoyer l'export", d:15, x:"Compter les lignes, supprimer les doublons, repérer les espaces parasites et les quantités en texte."},
      {k:'B', t:"Enrichir chaque ligne", d:25, x:"Colonnes G à M : mois, client propre, quantité, famille, commercial, prix de vente et d'achat."},
      {k:'C', t:"Calculer la rentabilité", d:15, x:"Colonnes N à T : CA brut, remise, CA net, coût, marge, taux, alerte « Marge faible »."},
      {k:'D', t:"Construire la synthèse", d:25, x:"Tableaux familles × mois, performance par commercial, zoom, top 5 clients, ligne de contrôle."},
      {k:'E', t:"Répondre au DAF", d:10, x:"Un message de 5 lignes : ce qui a été nettoyé, la cause chiffrée, ta proposition."}
    ]
  },
  corrige: `<p class="lead">La baisse de marge vient des remises de Lucas Moreau sur le petit électroménager depuis avril : son taux de marge sur cette famille passe de 25 % au 1er trimestre à 6,9 % en juin.</p>
  <div class="kpis"><div><span>CA net S1</span><b>454 124 €</b></div><div><span>Taux de marge global</span><b>31,6 %</b></div><div><span>Petit électroménager</span><b>23,1 %</b></div><div><span>PEM de Lucas, juin</span><b>6,9 %</b></div></div>
  <h4>A — Nettoyage</h4>
  <ul><li>1 402 lignes dans l'export, 2 doublons (FAC-2026-0256 et FAC-2026-0686), il reste 1 400 lignes.</li>
  <li>25 codes clients avec un espace final, 20 quantités stockées en texte.</li>
  <li>4 lignes à référence inconnue (PEM-099) : factures FAC-2026-0143, 0348, 0408 et 0680, à signaler au DAF.</li></ul>
  <h4>B et C — Formules en ligne 2</h4>
  <table class="plain f"><tbody>
  <tr><td>G</td><td><code>=MOIS(B2)</code></td></tr><tr><td>H</td><td><code>=SUPPRESPACE(C2)</code></td></tr><tr><td>I</td><td><code>=CNUM(E2)</code></td></tr>
  <tr><td>J</td><td><code>=SIERREUR(RECHERCHEV(D2;Produits!$A:$C;3;FAUX);"Réf inconnue")</code></td></tr>
  <tr><td>K</td><td><code>=RECHERCHEV(H2;Clients!$A:$E;5;FAUX)</code></td></tr>
  <tr><td>L</td><td><code>=SIERREUR(INDEX(Produits!$E:$E;EQUIV(D2;Produits!$A:$A;0));0)</code></td></tr>
  <tr><td>M</td><td><code>=SIERREUR(INDEX(Produits!$D:$D;EQUIV(D2;Produits!$A:$A;0));0)</code></td></tr>
  <tr><td>N à R</td><td><code>=I2*L2</code> · <code>=N2*F2</code> · <code>=N2-O2</code> · <code>=I2*M2</code> · <code>=P2-Q2</code></td></tr>
  <tr><td>S</td><td><code>=SI(P2=0;"";R2/P2)</code></td></tr>
  <tr><td>T</td><td><code>=SI(ET(P2>0;S2&lt;20%);"Marge faible";"")</code></td></tr></tbody></table>
  <h4>D — Synthèse</h4>
  <table class="plain f"><tbody>
  <tr><td>CA famille × mois</td><td><code>=SOMME.SI.ENS(Ventes!$P:$P;Ventes!$J:$J;$A5;Ventes!$G:$G;B$4)</code></td></tr>
  <tr><td>Remise moyenne</td><td><code>=MOYENNE.SI.ENS(Ventes!$F:$F;Ventes!$K:$K;$A29)</code></td></tr>
  <tr><td>Marges faibles</td><td><code>=NB.SI.ENS(Ventes!$K:$K;$A29;Ventes!$T:$T;"Marge faible")</code></td></tr>
  <tr><td>Top 5</td><td><code>=GRANDE.VALEUR(Clients!$F:$F;A41)</code> puis INDEX + EQUIV</td></tr>
  <tr><td>Contrôle</td><td><code>=SI(ARRONDI(H9-SOMME(Ventes!$P:$P);2)=0;"OK";"ÉCART")</code></td></tr></tbody></table>
  <div class="scroll"><table class="plain num"><thead><tr><th>Taux PEM</th><th>Janv.</th><th>Févr.</th><th>Mars</th><th>Avr.</th><th>Mai</th><th>Juin</th></tr></thead><tbody>
  <tr><td>Karim Benali</td><td>26,4 %</td><td>26,7 %</td><td>24,8 %</td><td>24,4 %</td><td>26,0 %</td><td>25,7 %</td></tr>
  <tr><td>Sophie Martin</td><td>24,3 %</td><td>26,9 %</td><td>24,4 %</td><td>26,7 %</td><td>26,7 %</td><td>23,5 %</td></tr>
  <tr class="bad"><td>Lucas Moreau</td><td>25,1 %</td><td>25,7 %</td><td>25,0 %</td><td>10,6 %</td><td>8,1 %</td><td>6,9 %</td></tr></tbody></table></div>
  <h4>E — L'analyse</h4>
  <p>Le taux global bouge peu (31,8 % en janvier, 31,1 % en juin) car les emballages ont doublé (24 336 € → 53 410 €) avec une marge stable de 33,5 % : c'est l'effet mix qui masque le problème. D'avril à juin, Lucas Moreau a accordé environ 4 170 € de remises sur 20 459 € de CA brut en petit électroménager (15 à 25 % au lieu de 0 à 10 %). Il ne lui reste que 1 445 € de marge sur 16 289 € de CA net.</p>
  <blockquote>Bonjour M. Faure, j'ai nettoyé l'export (2 doublons, 25 codes clients et 20 quantités mal formatés) et 4 ventes portent une référence inconnue (PEM-099) à vérifier. La hausse du CA vient surtout des emballages, qui ont doublé depuis janvier. La baisse de marge vient du petit électroménager chez Lucas Moreau : depuis avril, ses remises de 15 à 25 % ont fait tomber son taux de marge sur cette famille de 25 % à 7 % en juin. Je propose de faire le point avec lui et de fixer un plafond de remise. Je peux chiffrer l'impact d'un plafond à 10 % si vous le souhaitez.</blockquote>
  <p class="note">Le corrigé complet, avec toutes les formules dans les cellules, est dans les onglets orange du classeur.</p>`,
  autoeval: ["Je sais écrire RECHERCHEV sans regarder de tutoriel","Je sais expliquer pourquoi INDEX + EQUIV est plus souple","Je maîtrise les $ pour recopier une formule dans un tableau","Je sais utiliser SOMME.SI.ENS avec 2 ou 3 critères","J'ai trouvé seul la cause de la baisse de marge"],
  memo: [
    ["Références","$A$1 · $A1 · A$1","Bloquer colonne et/ou ligne ; touche F4"],
    ["Nettoyage","=SUPPRESPACE(C2)","Retire les espaces en trop"],
    ["Nettoyage","=CNUM(E2)","Convertit un nombre stocké en texte"],
    ["Nettoyage","=NBCAR(C2) · =ESTTEXTE(E2)","Repérer espaces parasites et textes"],
    ["Recherche","=RECHERCHEV(D2;Produits!$A:$C;3;FAUX)","Toujours FAUX pour un code"],
    ["Recherche","=INDEX(col_résultat;EQUIV(D2;col_recherche;0))","Plus robuste que RECHERCHEV"],
    ["Recherche","=SIERREUR(formule;\"message\")","Remplace une erreur par un message"],
    ["Logique","=SI(ET(A2>0;C2<20%);\"Alerte\";\"\")","ET : tous vrais · OU : au moins un"],
    ["Conditionnel","=SOMME.SI.ENS(somme;plage1;crit1;plage2;crit2)","La plage à additionner en premier"],
    ["Conditionnel","=NB.SI.ENS(plage1;crit1;…)","Compter les lignes"],
    ["Conditionnel","=MOYENNE.SI.ENS(moyenne;plage1;crit1)","Moyenne simple, non pondérée"],
    ["Classement","=GRANDE.VALEUR(plage;k)","k-ième plus grande valeur"],
    ["Gestion","Marge / prix d'achat","Taux de marge"],
    ["Gestion","Marge / prix de vente","Taux de marque"]
  ]
};
