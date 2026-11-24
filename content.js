/* Contenu de l'Atelier Excel & Power BI — Rhôna Distribution.
   Pour ajouter une séance : compléter PROGRAMME[n].ready = true et ajouter CONTENT.sXX. */

/* Grille au rendu Excel : zone Nom + barre de formule, en-têtes de lignes/colonnes, cellule active, références colorées, onglet.
   rows : lignes de valeurs. opts.cell : cellule active ("B4"), opts.fx : formule affichée, opts.hi : plage sélectionnée,
   opts.src / opts.refs : cellules utilisées par la formule (encadrées en couleur comme quand on appuie sur F2),
   opts.sheet : nom de l'onglet, opts.title : légende, opts.start : 1er n° de ligne, opts.col0 : décalage de la 1re lettre.
   Une valeur qui commence par ' est un nombre stocké en texte (aligné à gauche, triangle vert). */
function G(rows, opts) {
  opts = opts || {};
  const hi = opts.hi || [], start = opts.start || 1;
  const groups = opts.refs || (opts.src ? [opts.src] : []);
  const n = Math.max(...rows.map(r => r.length));
  const letters = Array.from({length: n}, (_, i) => String.fromCharCode(65 + (opts.col0 || 0) + i));
  const act = opts.cell && /^[A-Z]+\d+$/.test(opts.cell) ? opts.cell : null;
  const aCol = act ? act.replace(/\d+/, '') : null, aRow = act ? +act.replace(/[A-Z]+/, '') : null;
  const isNum = v => /^-?[\d  ]*\d([.,]\d+)?\s?(%|€| €)?$/.test(v);
  let h = '<figure class="xl">';
  if (opts.title) h += '<figcaption>' + opts.title + '</figcaption>';
  h += '<div class="xbar"><span class="xname">' + (act || '') + '</span><span class="xicons">✕ ✓ <i>fx</i></span><span class="xfx">' + (opts.fx || '') + '</span></div>';
  h += '<div class="scroll"><table><thead><tr><th class="xc"></th>' + letters.map(l => '<th' + (l === aCol ? ' class="xa"' : '') + '>' + l + '</th>').join('') + '</tr></thead><tbody>';
  rows.forEach((r, i) => {
    const rn = start + i;
    h += '<tr><th' + (rn === aRow ? ' class="xa"' : '') + '>' + rn + '</th>';
    for (let j = 0; j < n; j++) {
      const ref = letters[j] + rn;
      let v = r[j] === undefined ? '' : String(r[j]);
      const cls = [];
      if (v.startsWith("'")) { v = v.slice(1); cls.push('xt'); }
      else if (/^#/.test(v)) cls.push('xe');
      else if (v === 'VRAI' || v === 'FAUX') cls.push('xb');
      else if (v !== '' && isNum(v)) cls.push('xn');
      if (hi.includes(ref)) cls.push('xs');
      const g = groups.findIndex(gr => gr.includes(ref)); if (g >= 0) cls.push('xr' + (g % 4));
      if (ref === act) cls.push('xact');
      h += '<td class="' + cls.join(' ') + '">' + v + '</td>';
    }
    h += '</tr>';
  });
  h += '</tbody></table></div>';
  if (opts.sheet) h += '<div class="xtabs"><span>' + opts.sheet + '</span></div>';
  return h + '</figure>';
}
/* Exercice : énoncé (avec tableau) + solution masquée (avec tableau). */
const EXO = (n, titre, enonce, solution) => '<details class="exo"><summary><span class="exn">Exercice ' + n + '</span>' + titre + '</summary><div class="exb">' + enonce + '<details class="sol"><summary>Voir la solution</summary><div class="solb">' + solution + '</div></details></div></details>';
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
  `<p>Quand tu écris une formule puis que tu la recopies (en tirant la petite poignée verte en bas à droite de la cellule), Excel <b>ne recopie pas la formule à l'identique</b> : il décale automatiquement les cellules qu'elle utilise.</p>
  ${G([["Prix","Quantité","Total"],["10","3","30"],["25","2","50"],["8","5","40"]],{cell:"C3",fx:"=A3*B3",refs:[["A3"],["B3"]],hi:["C2","C3","C4"],sheet:"Exemple"})}
  <p>Exemple : en C2, on a écrit <code>=A2*B2</code>. Recopiée vers le bas, la formule devient <code>=A3*B3</code> en C3 (c'est ce qu'on voit dans la barre de formule ci-dessus), puis <code>=A4*B4</code> en C4. C'est pratique : chaque ligne calcule avec ses propres chiffres.</p>
  <p>Mais parfois, on veut qu'une cellule <b>ne bouge pas</b> pendant la recopie : un taux de TVA, un total, un paramètre. C'est le rôle du symbole <code>$</code>.</p>
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
  ${EXO(1,"La commission des commerciaux",
  `<p>Chaque commercial touche 3 % de ses ventes. Le taux de 3 % est écrit <b>une seule fois</b>, en E2. Tu veux calculer la commission en C2, puis recopier la formule vers le bas jusqu'en C4.</p>
  ${G([["Commercial","Ventes","Commission","","Taux"],["Karim","168 444","","","3%"],["Sophie","127 855","",""],["Lucas","157 825","",""]],{cell:"C2",sheet:"Commissions"})}
  <p><b>Question :</b> quelle formule écris-tu en C2 ?</p>`,
  `<p><b>Réponse :</b> <code>=B2*$E$2</code></p>
  ${G([["Commercial","Ventes","Commission","","Taux"],["Karim","168 444","5 053,32","","3%"],["Sophie","127 855","3 835,65",""],["Lucas","157 825","4 734,75",""]],{cell:"C4",fx:"=B4*$E$2",refs:[["B4"],["E2"]],hi:["C2","C3","C4"],sheet:"Commissions"})}
  <p>On voit ici la formule de la dernière ligne (C4) : <code>B2</code> est devenu <code>B4</code>, car chaque commercial a ses propres ventes. <code>$E$2</code> n'a pas bougé, car tout le monde utilise le même taux.</p>
  <p><b>Sans les $</b>, la formule de C4 serait <code>=B4*E4</code>. E4 est vide, donc le résultat serait 0.</p>`)}
  ${EXO(2,"Le prix après remise",
  `<p>Tu as 3 prix catalogue (colonne A) et 3 taux de remise (ligne 1). Tu veux remplir tout le tableau : chaque prix avec chaque remise. Tu écris une seule formule en B2, puis tu la recopies vers la droite <b>et</b> vers le bas.</p>
  ${G([["Prix \\ Remise","5%","10%","15%"],["100","","",""],["250","","",""],["80","","",""]],{cell:"B2",sheet:"Remises"})}
  <p><b>Rappel :</b> un prix après 5 % de remise = prix × (1 − 5 %).</p>
  <p><b>Question :</b> quelle formule écris-tu en B2 ?</p>`,
  `<p><b>Réponse :</b> <code>=$A2*(1-B$1)</code></p>
  ${G([["Prix \\ Remise","5%","10%","15%"],["100","95","90","85"],["250","237,5","225","212,5"],["80","76","72","68"]],{cell:"D4",fx:"=$A4*(1-D$1)",refs:[["A4"],["D1"]],hi:["B2","C2","D2","B3","C3","D3","B4","C4","D4"],sheet:"Remises"})}
  <p>Regarde la formule de la dernière cellule (D4) : <code>=$A4*(1-D$1)</code>.</p>
  <ul><li><code>$A</code> : le prix est toujours dans la colonne A, donc on bloque la colonne. La ligne, elle, change (2, 3, 4).</li>
  <li><code>$1</code> : la remise est toujours dans la ligne 1, donc on bloque la ligne. La colonne, elle, change (B, C, D).</li></ul>`)}
  ${EXO(3,"La part de chaque mois dans le total",
  `<p>Le CA de 3 mois est en B2:B4, et le total en B5. En C2, tu veux la part du mois de janvier dans le total, puis recopier la formule jusqu'en C4.</p>
  ${G([["Mois","CA net","Part"],["Janvier","64 663",""],["Février","62 486",""],["Mars","73 207",""],["Total","200 356",""]],{cell:"C2",sheet:"Parts"})}
  <p><b>Question :</b> quelle formule écris-tu en C2 ?</p>`,
  `<p><b>Réponse :</b> <code>=B2/$B$5</code></p>
  ${G([["Mois","CA net","Part"],["Janvier","64 663","32,3%"],["Février","62 486","31,2%"],["Mars","73 207","36,5%"],["Total","200 356",""]],{cell:"C3",fx:"=B3/$B$5",refs:[["B3"],["B5"]],hi:["C2","C3","C4"],sheet:"Parts"})}
  <p>La ligne du mois change à chaque fois (B2, B3, B4), mais on divise toujours par le même total (B5), donc on le bloque.</p>
  <p><b>Sans les $</b>, la formule de C3 serait <code>=B3/B6</code>. B6 est vide : Excel affiche l'erreur #DIV/0! (division par zéro).</p>`)}
  `},

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
  ${EXO(1,"Le total qui ne tombe pas juste",
  `<p>Tu additionnes 3 quantités : 10, 20 et 30. Le total devrait faire 60, mais Excel affiche 40.</p>
  ${G([["Quantité"],["10"],["'20"],["30"],["40"]],{cell:"A5",fx:"=SOMME(A2:A4)",refs:[["A2","A3","A4"]],sheet:"Ventes"})}
  <p>Regarde bien la cellule A3 : le 20 est collé à gauche, avec un petit triangle vert dans le coin.</p>
  <p><b>Question :</b> pourquoi le total est faux, et comment le corriger ?</p>`,
  `<p><b>Réponse :</b> le 20 est stocké <b>en texte</b>. SOMME ignore les textes, donc elle additionne seulement 10 + 30 = 40.</p>
  ${G([["Quantité","ESTTEXTE","Nombre"],["10","FAUX","10"],["'20","VRAI","20"],["30","FAUX","30"],["","","60"]],{cell:"C3",fx:"=CNUM(A3)",refs:[["A3"]],hi:["B3"],sheet:"Ventes"})}
  <ul><li>En colonne B, <code>=ESTTEXTE(A2)</code> le confirme : VRAI seulement pour le 20.</li>
  <li>En colonne C, <code>=CNUM(A2)</code> transforme chaque valeur en vrai nombre.</li>
  <li>En C5, <code>=SOMME(C2:C4)</code> donne enfin 60.</li></ul>`)}
  ${EXO(2,"Le client introuvable",
  `<p>Dans ton export, les codes clients doivent tous faire 6 caractères (exemple : CLI007). Tu calcules leur longueur avec NBCAR.</p>
  ${G([["Code client","Longueur"],["CLI007 ","7"],["CLI012","6"],["CLI003","6"]],{cell:"B2",fx:"=NBCAR(A2)",refs:[["A2"]],sheet:"Ventes"})}
  <p><b>Question :</b> quel code pose problème, et quelle formule permet de le nettoyer ?</p>`,
  `<p><b>Réponse :</b> CLI007, qui fait 7 caractères au lieu de 6 : il y a un espace invisible après le 7. On le nettoie avec <code>=SUPPRESPACE(A2)</code>.</p>
  ${G([["Code client","Longueur","Code propre","Longueur"],["CLI007 ","7","CLI007","6"],["CLI012","6","CLI012","6"],["CLI003","6","CLI003","6"]],{cell:"C2",fx:"=SUPPRESPACE(A2)",refs:[["A2"]],sheet:"Ventes"})}
  <p>Après nettoyage, tous les codes font 6 caractères : les recherches fonctionneront.</p>`)}
  `},

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
  <p>Pourquoi les <code>$</code> sur <code>Produits!$A:$C</code> ? Parce qu'on recopie la formule vers le bas sur 1 400 lignes : la table de recherche, elle, ne doit pas bouger (comme le taux de TVA unique de la notion 1 : il ne doit pas bouger).</p>
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
  ${EXO(1,"Retrouver le commercial du client",
  `<p>Tu as deux onglets. L'onglet <b>Clients</b> contient la liste des clients avec leur commercial :</p>
  ${G([["Code","Nom","Ville","Commercial"],["CLI001","Le Bouchon des Halles","Lyon 2e","Karim Benali"],["CLI002","Brasserie du Rhône","Lyon 3e","Sophie Martin"],["CLI003","Hôtel Bellecour","Lyon 6e","Lucas Moreau"]],{sheet:"Clients"})}
  <p>L'onglet <b>Ventes</b> contient tes factures, avec seulement le code client. Tu veux afficher le commercial en C2, puis recopier vers le bas.</p>
  ${G([["Facture","Code client","Commercial"],["FAC-0001","CLI003",""],["FAC-0002","CLI001",""]],{cell:"C2",sheet:"Ventes"})}
  <p><b>Question :</b> quelle formule écris-tu en C2 ?</p>`,
  `<p><b>Réponse :</b> <code>=RECHERCHEV(B2;Clients!$A:$D;4;FAUX)</code></p>
  ${G([["Facture","Code client","Commercial"],["FAC-0001","CLI003","Lucas Moreau"],["FAC-0002","CLI001","Karim Benali"]],{cell:"C2",fx:"=RECHERCHEV(B2;Clients!$A:$D;4;FAUX)",refs:[["B2"]],sheet:"Ventes"})}
  <ol><li><code>B2</code> : on cherche CLI003.</li>
  <li><code>Clients!$A:$D</code> : on cherche dans l'onglet Clients, colonnes A à D. Les codes sont bien dans la 1re colonne (A).</li>
  <li><code>4</code> : on compte les colonnes à partir de A : A = 1, B = 2, C = 3, D = 4. Le commercial est dans la 4e.</li>
  <li><code>FAUX</code> : on veut exactement CLI003, pas un code « proche ».</li></ol>`)}
  `},

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
  ${EXO(1,"Chercher vers la gauche",
  `<p>Dans ce tableau, le nom est en colonne A et le code en colonne B. Tu connais le <b>code</b> d'un client (écrit en D2) et tu veux retrouver son <b>nom</b> en E2.</p>
  ${G([["Nom","Code","","Code cherché","Nom trouvé"],["Le Bouchon des Halles","CLI001","","CLI003",""],["Brasserie du Rhône","CLI002"],["Hôtel Bellecour","CLI003"],["Café des Terreaux","CLI004"]],{cell:"E2",sheet:"Clients"})}
  <p><b>Question :</b> pourquoi RECHERCHEV ne peut pas faire ça, et quelle formule utiliser en E2 ?</p>`,
  `<p><b>Pourquoi pas RECHERCHEV :</b> elle cherche toujours dans la 1re colonne de la plage, puis renvoie une colonne <b>à droite</b>. Ici, le code (ce qu'on cherche) est à droite du nom (ce qu'on veut) : impossible.</p>
  <p><b>Réponse :</b> <code>=INDEX(A:A;EQUIV(D2;B:B;0))</code></p>
  ${G([["Nom","Code","","Code cherché","Nom trouvé"],["Le Bouchon des Halles","CLI001","","CLI003","Hôtel Bellecour"],["Brasserie du Rhône","CLI002"],["Hôtel Bellecour","CLI003"],["Café des Terreaux","CLI004"]],{cell:"E2",fx:"=INDEX(A:A;EQUIV(D2;B:B;0))",refs:[["D2"],["B4"],["A4"]],sheet:"Clients"})}
  <ol><li><code>EQUIV(D2;B:B;0)</code> cherche CLI003 dans la colonne B : il est à la ligne <b>4</b>. EQUIV renvoie 4.</li>
  <li><code>INDEX(A:A;4)</code> va dans la colonne A et prend la 4e cellule : <b>Hôtel Bellecour</b>.</li></ol>`)}
  `},

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
  ${EXO(1,"La prime des commerciaux",
  `<p>Un commercial touche une prime de 500 € seulement s'il remplit <b>les deux</b> conditions : son CA dépasse 50 000 €, <b>et</b> son taux de marge est d'au moins 30 %. Sinon, il touche 0 €.</p>
  ${G([["Commercial","CA","Taux de marge","Prime"],["Karim","62 000","32%",""],["Sophie","48 000","35%",""],["Lucas","71 000","18%",""]],{cell:"D2",sheet:"Primes"})}
  <p><b>Question :</b> quelle formule écris-tu en D2, pour la recopier jusqu'en D4 ?</p>`,
  `<p><b>Réponse :</b> <code>=SI(ET(B2>50000;C2>=30%);500;0)</code></p>
  ${G([["Commercial","CA","Taux de marge","Prime"],["Karim","62 000","32%","500"],["Sophie","48 000","35%","0"],["Lucas","71 000","18%","0"]],{cell:"D2",fx:"=SI(ET(B2>50000;C2>=30%);500;0)",refs:[["B2"],["C2"]],hi:["D2","D3","D4"],sheet:"Primes"})}
  <ul><li><b>Karim</b> : 62 000 &gt; 50 000 ✓ et 32 % ≥ 30 % ✓ → 500 €.</li>
  <li><b>Sophie</b> : 48 000 n'est pas &gt; 50 000 ✗ → 0 €, même si sa marge est bonne.</li>
  <li><b>Lucas</b> : CA ✓ mais 18 % &lt; 30 % ✗ → 0 €.</li></ul>
  <p>Avec <code>OU</code> à la place de <code>ET</code>, il suffirait d'une seule condition : les trois toucheraient la prime.</p>`)}
  `},

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
  <p>On combine SOMME.SI.ENS et la technique du tableau à double entrée vue dans la notion 1 : les familles sont écrites en colonne A (une par ligne), les numéros de mois en ligne 4 (un par colonne), et une seule formule remplit tout le tableau.</p>
  ${G([["Famille \\ Mois","1","2","3"],["Vaisselle","11 795","7 950","12 323"],["Emballages","24 336","26 217","29 524"]],{start:4,fx:"=SOMME.SI.ENS(Ventes!$P:$P;Ventes!$J:$J;$A5;Ventes!$G:$G;B$4)",cell:"B5",hi:["B5","C5","D5","B6","C6","D6"],src:["A5","A6","B4","C4","D4"]})}
  <ul><li><code>Ventes!$P:$P</code>, <code>$J:$J</code>, <code>$G:$G</code> : colonnes de données, totalement bloquées.</li>
  <li><code>$A5</code> : la famille, toujours en colonne A, mais une par ligne.</li>
  <li><code>B$4</code> : le mois, toujours en ligne 4, mais un par colonne.</li></ul>
  ${TRAP("Toujours une ligne de contrôle","Le total de ton tableau de synthèse doit être égal à SOMME de la colonne source. Si une famille est mal orthographiée en colonne A, ses ventes disparaissent du tableau sans erreur : seul le contrôle te le dira.")}
  <h4>À toi de jouer</h4>
  <p>Les deux exercices utilisent ce petit tableau de ventes :</p>
  ${G([["Commercial","Mois","CA net"],["Lucas Moreau","4","1 352"],["Karim Benali","4","215"],["Lucas Moreau","5","598"],["Sophie Martin","5","1 240"],["Lucas Moreau","6","2 100"],["Karim Benali","6","980"]],{sheet:"Ventes"})}
  ${EXO(1,"Compter les grosses ventes",
  `<p>Combien de ventes dépassent 1 000 € de CA net ? Tu écris la formule en E2.</p>`,
  `<p><b>Réponse :</b> <code>=NB.SI.ENS(C2:C7;">1000")</code> → <b>3</b></p>
  ${G([["Commercial","Mois","CA net","","Nb > 1 000 €"],["Lucas Moreau","4","1 352","","3"],["Karim Benali","4","215"],["Lucas Moreau","5","598"],["Sophie Martin","5","1 240"],["Lucas Moreau","6","2 100"],["Karim Benali","6","980"]],{cell:"E2",fx:'=NB.SI.ENS(C2:C7;">1000")',refs:[["C2","C5","C6"]],sheet:"Ventes"})}
  <p>Les 3 ventes comptées sont encadrées : 1 352, 1 240 et 2 100. Le critère <code>">1000"</code> s'écrit entre guillemets.</p>`)}
  ${EXO(2,"Le CA de Lucas en mai et juin",
  `<p>Tu veux le CA net de Lucas Moreau pour les mois 5 et 6 seulement. Tu écris la formule en E2.</p>`,
  `<p><b>Réponse :</b> <code>=SOMME.SI.ENS(C2:C7;A2:A7;"Lucas Moreau";B2:B7;">=5")</code> → <b>2 698</b></p>
  ${G([["Commercial","Mois","CA net","","CA Lucas mai-juin"],["Lucas Moreau","4","1 352","","2 698"],["Karim Benali","4","215"],["Lucas Moreau","5","598"],["Sophie Martin","5","1 240"],["Lucas Moreau","6","2 100"],["Karim Benali","6","980"]],{cell:"E2",fx:'=SOMME.SI.ENS(C2:C7;A2:A7;"Lucas Moreau";B2:B7;">=5")',refs:[["C4","C6"]],sheet:"Ventes"})}
  <ul><li>1re condition : le commercial est Lucas Moreau (lignes 2, 4 et 6).</li>
  <li>2e condition : le mois est supérieur ou égal à 5 (lignes 4 et 6).</li>
  <li>Seules les lignes 4 et 6 remplissent les deux conditions : 598 + 2 100 = 2 698.</li></ul>
  <p>Pour une fourchette fermée (par exemple mois 4 à 5), on met deux critères sur la même colonne : <code>B2:B7;">=4";B2:B7;"<=5"</code>.</p>`)}
  `},

  {id:'top', t:"7. Faire un classement : GRANDE.VALEUR", d:10, html:
  `${F("=GRANDE.VALEUR(plage ; k)    → la k-ième plus grande valeur\n=PETITE.VALEUR(plage ; k)    → la k-ième plus petite")}
  <p><code>=GRANDE.VALEUR(F2:F31;1)</code> donne le plus gros CA, <code>;2</code> le 2e, etc. Pour un top 5, on écrit les rangs 1 à 5 dans une colonne et on s'en sert comme k.</p>
  <p>GRANDE.VALEUR donne un montant, pas un nom. Pour retrouver le client qui correspond, on réutilise INDEX + EQUIV (notion 4) : EQUIV cherche ce montant dans la colonne des CA, INDEX renvoie le nom à la même position.</p>
  ${G([["Rang","Client","CA net"],["1","Restaurant L'Écullois","32 473"],["2","Restaurant Le Guillotière","26 837"]],{fx:"=INDEX(Clients!$B:$B;EQUIV(C2;Clients!$F:$F;0))",cell:"B2",hi:["B2"],src:["A2","C2"]})}
  <p>En C2 : <code>=GRANDE.VALEUR(Clients!$F:$F;A2)</code>. En B2 : la formule ci-dessus.</p>
  ${TRAP("Limite à connaître","Si deux clients ont exactement le même CA, EQUIV renverra deux fois le premier. Avec des montants au centime près, c'est très rare, mais il faut le savoir.")}
  <h4>À toi de jouer</h4>
  ${EXO(1,"Le 1er et le 2e client",
  `<p>Voici le CA de 5 clients. En E2 et E3, tu veux les deux plus gros CA, et en F2 et F3 le nom du client correspondant.</p>
  ${G([["Client","CA net","","Rang","CA","Client"],["Le Bouchon","18 200","","1","",""],["La Brasserie","25 400","","2","",""],["L'Hôtel","31 900"],["Le Café","12 300"],["Le Traiteur","22 750"]],{cell:"E2",sheet:"Top clients"})}
  <p><b>Question :</b> quelles formules écris-tu en E2 et en F2 (à recopier en ligne 3) ?</p>`,
  `<p><b>Réponse :</b> E2 <code>=GRANDE.VALEUR($B$2:$B$6;D2)</code> · F2 <code>=INDEX($A$2:$A$6;EQUIV(E2;$B$2:$B$6;0))</code></p>
  ${G([["Client","CA net","","Rang","CA","Client"],["Le Bouchon","18 200","","1","31 900","L'Hôtel"],["La Brasserie","25 400","","2","25 400","La Brasserie"],["L'Hôtel","31 900"],["Le Café","12 300"],["Le Traiteur","22 750"]],{cell:"F2",fx:"=INDEX($A$2:$A$6;EQUIV(E2;$B$2:$B$6;0))",refs:[["E2"],["B4"],["A4"]],sheet:"Top clients"})}
  <ol><li>E2 : le rang 1 (D2) donne le plus grand CA, 31 900. En E3, D3 vaut 2 : on obtient le 2e, 25 400.</li>
  <li>F2 : EQUIV cherche 31 900 dans B2:B6 et le trouve en 3e position. INDEX prend la 3e valeur de A2:A6 : L'Hôtel.</li>
  <li>Les plages sont bloquées avec des $ parce qu'on recopie vers le bas.</li></ol>`)}
  `},

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
  ${EXO(1,"Calculer la marge et la remise",
  `<p>Un produit est acheté 40 € et vendu 60 €. Ligne 3, le même produit est vendu avec 10 % de remise, donc à 54 €.</p>
  ${G([["","Prix d'achat","Prix de vente","Marge","Taux de marge","Taux de marque"],["Sans remise","40","60","","",""],["Remise 10 %","40","54","","",""]],{cell:"D2",sheet:"Marges"})}
  <p><b>Question :</b> quelles formules écris-tu en D2, E2 et F2, et quels résultats obtiens-tu sur les deux lignes ?</p>`,
  `<p><b>Réponse :</b> D2 <code>=C2-B2</code> · E2 <code>=D2/B2</code> · F2 <code>=D2/C2</code>, puis on recopie sur la ligne 3.</p>
  ${G([["","Prix d'achat","Prix de vente","Marge","Taux de marge","Taux de marque"],["Sans remise","40","60","20","50,0%","33,3%"],["Remise 10 %","40","54","14","35,0%","25,9%"]],{cell:"F3",fx:"=D3/C3",refs:[["D3"],["C3"]],hi:["D2","E2","F2","D3","E3","F3"],sheet:"Marges"})}
  <p>La remise de 10 % fait passer la marge de 20 € à 14 € : on perd 30 % de la marge, pas 10 %. C'est exactement le problème de Lucas dans le cas pratique.</p>`)}
  `},

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
    {q:"Le taux de TVA est en B1. En B4, on a écrit =A4*B1, puis on a recopié la formule vers le bas. B4 est juste, mais B5 et B6 affichent 0. Quelle formule fallait-il écrire en B4 ?",
     h:G([["Taux TVA","20%"],["",""],["Prix HT","TVA"],["100","20"],["250","0"],["80","0"]],{cell:"B5",fx:"=A5*B2",refs:[["A5"],["B2"]],sheet:"Quiz 1"}),
     o:["=A4*B1","=$A$4*B1","=A4*$B$1","=A$4*B1"], a:2,
     e:"En B5, la formule est devenue =A5*B2 (on le voit dans la barre de formule) : B2 est vide, d'où le 0. Il faut bloquer le taux : =A4*$B$1. A4 reste sans $ pour que chaque ligne prenne son propre prix."},
    {q:"On veut remplir ce tableau (familles × mois) avec une seule formule en B5, recopiée vers la droite et vers le bas. Comment écrire la référence à la famille, qui est en A5 ?",
     h:G([["Famille \\ Mois","1","2","3"],["Vaisselle","","",""],["Emballages","","",""]],{start:4,cell:"B5",refs:[["A5","A6"],["B4","C4","D4"]],sheet:"Quiz 2"}),
     o:["A5","$A5","A$5","$A$5"], a:1,
     e:"La famille est toujours en colonne A (on bloque la colonne avec $A), mais elle change à chaque ligne (on laisse la ligne libre). Donc $A5. Le mois, lui, s'écrira B$4."},
    {q:"Dans l'onglet Ventes, RECHERCHEV renvoie #N/A pour le client de la ligne 2, alors que CLI007 existe bien dans l'onglet Clients. Quelle est la cause la plus probable ?",
     h:G([["Code client","Longueur","Commercial"],["CLI007 ","7","#N/A"],["CLI012","6","Sophie Martin"]],{cell:"C2",fx:"=RECHERCHEV(A2;Clients!$A:$E;5;FAUX)",refs:[["A2"]],sheet:"Quiz 3"}),
     o:["Le 4e argument FAUX est une erreur","Il y a un espace invisible après le code","La table Clients est trop grande","Il manque un $ devant A2"], a:1,
     e:"La colonne Longueur (=NBCAR) indique 7 caractères au lieu de 6 : il y a un espace après CLI007. Pour Excel, « CLI007 » avec un espace et « CLI007 » sont deux textes différents. SUPPRESPACE règle le problème."},
    {q:"Quelle formule additionne le CA de Lucas Moreau dans ce tableau ?",
     h:G([["Commercial","Mois","CA net"],["Lucas Moreau","4","1 352"],["Karim Benali","4","215"],["Lucas Moreau","5","598"]],{refs:[["C2","C4"]],sheet:"Quiz 4"}),
     o:['=SOMME.SI.ENS(A:A;"Lucas Moreau";C:C)','=SOMME.SI.ENS(C:C;A:A;"Lucas Moreau")','=SOMME.SI.ENS("Lucas Moreau";A:A;C:C)','=SOMME.SI.ENS(C:C;"Lucas Moreau";A:A)'], a:1,
     e:"Avec SOMME.SI.ENS, la colonne à additionner (C, le CA) vient en premier. Ensuite on donne la colonne du critère (A) puis le critère (\"Lucas Moreau\"). Résultat : 1 352 + 598 = 1 950."},
    {q:"Ce produit est acheté 100 € et vendu 125 €. Quel est son taux de marque ?",
     h:G([["Prix d'achat","Prix de vente","Marge"],["100","125","25"]],{sheet:"Quiz 5"}),
     o:["25 %","20 %","125 %","80 %"], a:1,
     e:"Taux de marque = marge / prix de vente = 25 / 125 = 20 %. Le taux de marge, lui, se calcule sur le prix d'achat : 25 / 100 = 25 %."},
    {q:"Le nom du client est en colonne A, son code en colonne B. On connaît le code (en D2) et on veut le nom. Quelle formule fonctionne ?",
     h:G([["Nom","Code","","Code cherché"],["Le Bouchon","CLI001","","CLI002"],["La Brasserie","CLI002"],["L'Hôtel","CLI003"]],{cell:"E2",refs:[["D2"]],sheet:"Quiz 6"}),
     o:["=RECHERCHEV(D2;A:B;1;FAUX)","=INDEX(A:A;EQUIV(D2;B:B;0))","=EQUIV(D2;A:A;0)","=INDEX(B:B;EQUIV(D2;A:A;0))"], a:1,
     e:"EQUIV cherche CLI002 dans la colonne B des codes et trouve la position 3. INDEX prend la 3e cellule de la colonne A des noms : La Brasserie. RECHERCHEV ne marche pas ici, car elle ne peut chercher que dans la 1re colonne de la plage (A, les noms)."}
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
