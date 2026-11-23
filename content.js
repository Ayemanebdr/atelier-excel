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
  {id:'refs', t:"1. Les références : relatives, absolues, mixtes", d:15, html:
  `<p>Quand tu recopies une formule, Excel décale automatiquement les cellules qu'elle utilise. C'est souvent ce qu'on veut, mais pas toujours. Le symbole <code>$</code> sert à <b>bloquer</b> une colonne, une ligne, ou les deux.</p>
  <table class="plain"><thead><tr><th>Écriture</th><th>Nom</th><th>Quand tu recopies…</th></tr></thead><tbody>
  <tr><td><code>A1</code></td><td>Relative</td><td>la colonne et la ligne bougent</td></tr>
  <tr><td><code>$A$1</code></td><td>Absolue</td><td>rien ne bouge</td></tr>
  <tr><td><code>$A1</code></td><td>Mixte</td><td>la colonne A reste fixe, la ligne bouge</td></tr>
  <tr><td><code>A$1</code></td><td>Mixte</td><td>la ligne 1 reste fixe, la colonne bouge</td></tr></tbody></table>
  <h4>Exemple : appliquer un taux de TVA</h4>
  <p>Le taux est dans une seule cellule (B1). Si tu écris <code>=A4*B1</code> et que tu recopies vers le bas, Excel écrit <code>=A5*B2</code>, puis <code>=A6*B3</code>… et B2, B3 sont vides. Résultat : des zéros. Il faut bloquer B1.</p>
  ${G([["Taux TVA","20%",""],["","",""],["Prix HT","TVA",""],["100","20",""],["250","50",""],["80","16",""]],{fx:"=A4*$B$1",cell:"B4",hi:["B4","B5","B6"],src:["B1"]})}
  <h4>Exemple : un tableau à double entrée (la technique clé de la séance)</h4>
  <p>Tu veux remplir un tableau « familles × mois ». La famille est toujours en colonne A, le mois toujours en ligne 1. Une seule formule suffit si tu bloques <b>la colonne A</b> (<code>$A2</code>) et <b>la ligne 1</b> (<code>B$1</code>).</p>
  ${G([["Famille \\ Mois","1","2","3"],["Vaisselle","11 795","7 950","12 323"],["Emballages","24 336","26 217","29 524"]],{fx:"=SOMME.SI.ENS(Ventes!$P:$P;Ventes!$J:$J;$A2;Ventes!$G:$G;B$1)",cell:"B2",hi:["B2","C2","D2","B3","C3","D3"],src:["A2","A3","B1","C1","D1"]})}
  ${TIP("Astuce : la touche F4","Clique dans la formule sur une référence et appuie sur F4 : A1 → $A$1 → A$1 → $A1 → A1. Plus besoin de taper les $ à la main.")}
  ${TRAP("Piège classique","Un tableau qui affiche des zéros ou des chiffres absurdes après recopie : c'est presque toujours un $ oublié. Clique sur une cellule du bas du tableau et lis sa formule : tu verras tout de suite quelle référence a glissé.")}`},

  {id:'clean', t:"2. Nettoyer un export avant de l'analyser", d:10, html:
  `<p>En entreprise, les données sortent d'un logiciel (facturation, comptabilité, ERP) et ne sont <b>jamais</b> parfaites. Trois défauts reviennent tout le temps, et ils sont dangereux parce qu'ils ne se voient pas.</p>
  <table class="plain"><thead><tr><th>Défaut</th><th>Comment le repérer</th><th>Comment le corriger</th></tr></thead><tbody>
  <tr><td>Espaces en trop (<code>"CLI007 "</code>)</td><td><code>=NBCAR(C2)</code> donne 7 au lieu de 6</td><td><code>=SUPPRESPACE(C2)</code></td></tr>
  <tr><td>Nombres stockés en texte</td><td><code>=ESTTEXTE(E2)</code> renvoie VRAI ; nombre aligné à gauche, triangle vert</td><td><code>=CNUM(E2)</code></td></tr>
  <tr><td>Lignes en double</td><td>Mise en forme conditionnelle &gt; Valeurs en double</td><td>Données &gt; Supprimer les doublons</td></tr></tbody></table>
  ${G([["Code brut","NBCAR","Propre"],["CLI007 ","7","CLI007"],["CLI012","6","CLI012"]],{fx:"=SUPPRESPACE(A2)",cell:"C2",hi:["C2"],src:["A2"]})}
  <p><b>Pourquoi c'est grave :</b> RECHERCHEV ne trouve pas <code>"CLI007 "</code> dans la liste des clients (il cherche le texte exact, espace compris) et renvoie #N/A. Et SOMME <b>ignore</b> les nombres en texte : ton total est faux, sans aucun message d'erreur.</p>
  ${TIP("Bonne pratique","Ne modifie pas les colonnes d'origine : ajoute une colonne « propre » à côté avec la formule. Tu gardes une trace de ce qui a été corrigé, et un contrôleur peut toujours vérifier.")}`},

  {id:'rechv', t:"3. RECHERCHEV : aller chercher une info dans une autre table", d:15, html:
  `<p>Ton export de ventes contient seulement la référence produit (<code>VAI-001</code>). Le prix et la famille sont dans une autre table, l'onglet Produits. RECHERCHEV va chercher l'information pour toi.</p>
  ${F("=RECHERCHEV(valeur_cherchée ; table ; n°_colonne ; FAUX)")}
  <ol><li><b>valeur_cherchée</b> : ce que tu cherches (la référence, en D2).</li>
  <li><b>table</b> : où chercher. La valeur cherchée doit être dans la <b>première colonne</b> de cette plage.</li>
  <li><b>n°_colonne</b> : quelle colonne renvoyer, en comptant depuis la première colonne de la table (1, 2, 3…).</li>
  <li><b>FAUX</b> : correspondance exacte. Toujours FAUX (ou 0) pour chercher un code.</li></ol>
  ${G([["Réf produit","Désignation","Famille","Prix d'achat","Prix de vente"],["VAI-001","Assiette plate (lot de 6)","Vaisselle","14,20","23,90"],["PEM-002","Blender pro 2 L","Petit électroménager","215,00","299,00"],["EMB-001","Boîte repas kraft","Emballages","38,50","59,00"]],{title:"Onglet Produits",src:["A3"],hi:["C3"]})}
  <p>Pour une vente de <code>PEM-002</code> en D2, la famille s'obtient avec :</p>
  ${F("=RECHERCHEV(D2;Produits!$A:$C;3;FAUX)   →   Petit électroménager")}
  <p>Excel cherche PEM-002 dans la colonne A de Produits, trouve la ligne 3, et renvoie la 3e colonne de la plage A:C, donc C3.</p>
  ${TRAP("Le piège n°1 en entreprise","Oublier FAUX. Sans lui, Excel fait une recherche « approchée » : il renvoie une valeur, mais pas forcément la bonne, et sans aucune erreur. Des reportings entiers sont faux à cause de ça.")}
  <h4>Lire les erreurs</h4>
  <table class="plain"><tbody><tr><td><code>#N/A</code></td><td>Valeur introuvable : code inexistant, espace parasite, ou nombre stocké en texte d'un côté et en nombre de l'autre.</td></tr>
  <tr><td><code>#REF!</code></td><td>Le n° de colonne dépasse la largeur de la table (ex. 5 avec une plage A:C).</td></tr></tbody></table>
  <h4>SIERREUR : remplacer l'erreur par un message utile</h4>
  ${F('=SIERREUR(RECHERCHEV(D2;Produits!$A:$C;3;FAUX);"Réf inconnue")')}
  ${TIP("Astuce","Utilise SIERREUR après avoir compris tes erreurs, pas avant : s'il cache tout dès le début, tu ne vois plus les vrais problèmes de données. Ici, « Réf inconnue » est un message qui t'aide à les repérer.")}`},

  {id:'index', t:"4. INDEX + EQUIV : la version professionnelle", d:15, html:
  `<p>RECHERCHEV a deux limites : la valeur cherchée doit être dans la première colonne, et le n° de colonne est « en dur ». Si quelqu'un insère une colonne dans la table, ta formule renvoie la mauvaise donnée. INDEX + EQUIV n'a pas ces défauts.</p>
  ${F("=EQUIV(valeur ; colonne_où_chercher ; 0)     → renvoie la POSITION (n° de ligne)\n=INDEX(colonne_à_renvoyer ; position)         → renvoie la VALEUR à cette position")}
  <p>On les combine : EQUIV trouve la ligne, INDEX va chercher la valeur dans la colonne voulue.</p>
  ${F("=INDEX(Produits!$E:$E;EQUIV(D2;Produits!$A:$A;0))   →   299,00")}
  ${G([["Réf produit","…","…","Prix d'achat","Prix de vente"],["VAI-001","","","14,20","23,90"],["PEM-002","","","215,00","299,00"]],{title:"EQUIV trouve PEM-002 en position 3 de la colonne A, INDEX renvoie la position 3 de la colonne E",src:["A3"],hi:["E3"]})}
  <table class="plain"><thead><tr><th></th><th>RECHERCHEV</th><th>INDEX + EQUIV</th></tr></thead><tbody>
  <tr><td>Colonne de recherche</td><td>Forcément la 1re</td><td>N'importe laquelle</td></tr>
  <tr><td>Insertion d'une colonne</td><td>Casse la formule</td><td>Aucun problème</td></tr>
  <tr><td>Facilité</td><td>Plus simple à écrire</td><td>Un peu plus long</td></tr></tbody></table>
  ${TIP("Pour l'entretien","Depuis Excel 365, RECHERCHEX remplace les deux : =RECHERCHEX(D2;Produits!A:A;Produits!E:E). Ton Excel 2013 ne l'a pas, mais savoir qu'elle existe et ce qu'elle fait est un plus. INDEX + EQUIV marche partout, c'est pour ça qu'on l'apprend d'abord.")}`},

  {id:'si', t:"5. SI, ET, OU : faire réagir le tableau", d:10, html:
  `${F('=SI(test ; valeur_si_vrai ; valeur_si_faux)')}
  <p>Exemple : signaler les ventes dont le taux de marge est sous 20 %.</p>
  ${G([["CA net","Marge","Taux","Alerte"],["1 690","169","10,0%","Marge faible"],["299","84","28,1%",""],["0","0","",""]],{fx:'=SI(ET(A2>0;C2<20%);"Marge faible";"")',cell:"D2",hi:["D2"]})}
  <ul><li><code>ET(test1;test2)</code> est VRAI si <b>tous</b> les tests sont vrais.</li>
  <li><code>OU(test1;test2)</code> est VRAI si <b>au moins un</b> test est vrai.</li>
  <li><code>""</code> (deux guillemets) affiche une cellule vide.</li></ul>
  <h4>Éviter la division par zéro</h4>
  ${F('=SI(A2=0;"";B2/A2)')}
  <p>Sans ce SI, une ligne à CA nul afficherait #DIV/0!, et cette erreur se propagerait dans tous tes totaux.</p>`},

  {id:'sommesi', t:"6. SOMME.SI.ENS, NB.SI.ENS, MOYENNE.SI.ENS", d:20, html:
  `<p>Ce sont les fonctions les plus utilisées en contrôle de gestion : additionner, compter ou faire une moyenne <b>seulement pour les lignes qui respectent des critères</b>.</p>
  ${F("=SOMME.SI.ENS(plage_à_additionner ; plage_critère1 ; critère1 ; plage_critère2 ; critère2 ; …)")}
  ${G([["Commercial","Famille","Mois","CA net"],["Lucas Moreau","Petit électroménager","4","1 352"],["Karim Benali","Vaisselle","4","215"],["Lucas Moreau","Petit électroménager","5","598"],["Lucas Moreau","Emballages","5","472"]],{title:"Combien Lucas a-t-il vendu en petit électroménager ?",hi:["D2","D4"]})}
  ${F('=SOMME.SI.ENS(D:D;A:A;"Lucas Moreau";B:B;"Petit électroménager")   →   1 950')}
  <table class="plain"><thead><tr><th>Fonction</th><th>Ce qu'elle fait</th><th>Particularité</th></tr></thead><tbody>
  <tr><td><code>SOMME.SI.ENS</code></td><td>Additionne</td><td>La plage à additionner vient <b>en premier</b></td></tr>
  <tr><td><code>NB.SI.ENS</code></td><td>Compte les lignes</td><td>Pas de plage à additionner, que des critères</td></tr>
  <tr><td><code>MOYENNE.SI.ENS</code></td><td>Fait la moyenne</td><td>Même ordre que SOMME.SI.ENS</td></tr></tbody></table>
  <h4>Écrire les critères</h4>
  <ul><li>Texte : <code>"Vaisselle"</code> ou mieux, une cellule : <code>$A5</code>.</li>
  <li>Comparaison : <code>"&gt;1000"</code>, <code>"&lt;&gt;0"</code> (différent de 0).</li>
  <li>Comparaison avec une cellule : <code>"&gt;="&amp;B1</code>.</li></ul>
  ${TIP("La méthode du contrôleur","Ne tape jamais le critère en dur dans 24 formules. Mets les familles en colonne A, les mois en ligne d'en-tête, et écris une seule formule avec $A5 et B$4 (voir la notion 1). Si demain on ajoute une famille, tu ajoutes une ligne et tu recopies.")}
  ${TRAP("Piège","Ajoute toujours une ligne de contrôle : le total de ton tableau de synthèse doit être égal à SOMME de la colonne source. Si une famille est mal orthographiée dans ta colonne A, ses ventes disparaissent du tableau sans erreur, et seul le contrôle te le dira.")}`},

  {id:'top', t:"7. Faire un classement : GRANDE.VALEUR", d:5, html:
  `${F("=GRANDE.VALEUR(plage ; k)    → la k-ième plus grande valeur\n=PETITE.VALEUR(plage ; k)    → la k-ième plus petite")}
  <p>Pour un top 5 clients : calcule d'abord le CA de chaque client (avec SOMME.SI.ENS), puis :</p>
  ${G([["Rang","Client","CA net"],["1","Restaurant L'Écullois","32 473"],["2","Restaurant Le Guillotière","26 837"]],{fx:"=INDEX(Clients!$B:$B;EQUIV(C2;Clients!$F:$F;0))",cell:"B2",hi:["B2"],src:["A2","C2"]})}
  <p>C2 contient <code>=GRANDE.VALEUR(Clients!$F:$F;A2)</code> et B2 retrouve le nom du client avec INDEX + EQUIV : tu réutilises la notion 4.</p>`},

  {id:'gestion', t:"8. Les notions de gestion derrière les formules", d:15, html:
  `<p>Excel n'est qu'un outil : ce qu'on attend de toi en stage, c'est de comprendre ce que tu calcules.</p>
  <table class="plain"><tbody>
  <tr><td><b>CA brut</b></td><td>Quantité × prix de vente catalogue</td></tr>
  <tr><td><b>Remise</b></td><td>Réduction accordée au client (en % ou en €)</td></tr>
  <tr><td><b>CA net</b></td><td>CA brut − remises : ce que l'entreprise facture vraiment</td></tr>
  <tr><td><b>Coût d'achat</b></td><td>Quantité × prix d'achat fournisseur</td></tr>
  <tr><td><b>Marge commerciale</b></td><td>CA net − coût d'achat</td></tr></tbody></table>
  <h4>Taux de marge ou taux de marque ? La question qui tombe en entretien</h4>
  ${G([["Prix d'achat","Prix de vente","Marge","Taux de marge","Taux de marque"],["100","125","25","25 %","20 %"]],{hi:["D2","E2"]})}
  <ul><li><b>Taux de marge</b> = marge / prix d'achat = 25 / 100 = 25 %</li>
  <li><b>Taux de marque</b> = marge / prix de vente = 25 / 125 = 20 %</li></ul>
  <p>Dans beaucoup d'entreprises, on dit « taux de marge » pour marge / CA (en réalité le taux de marque). Dans ce cas pratique, on suit cet usage. Mais montrer que tu connais la nuance fait bonne impression.</p>
  <h4>Pourquoi une remise coûte si cher</h4>
  <p>Un produit acheté 215 € et vendu 299 € rapporte 84 € (28 % du prix). Avec 20 % de remise, il est vendu 239 € : il ne rapporte plus que 24 €. <b>20 % de remise = 71 % de la marge perdue.</b> C'est exactement ce que tu vas découvrir dans le cas.</p>
  ${G([["","Sans remise","Remise 20 %"],["Prix de vente","299","239"],["Prix d'achat","215","215"],["Marge","84","24"]],{hi:["C4"]})}
  <h4>L'effet mix</h4>
  <p>Un taux de marge global peut rester stable alors qu'une famille s'effondre, si une autre famille à bonne marge progresse en même temps. C'est pour ça qu'un contrôleur <b>découpe toujours</b> : par famille, puis par commercial, puis par mois.</p>`},

  {id:'raccourcis', t:"9. Les raccourcis qui font gagner du temps", d:5, html:
  `<table class="plain"><tbody>
  <tr><td><kbd>Ctrl</kbd> + <kbd>↓</kbd></td><td>Aller à la dernière ligne remplie</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>Maj</kbd> + <kbd>↓</kbd></td><td>Sélectionner jusqu'à la dernière ligne</td></tr>
  <tr><td>Double-clic sur la poignée</td><td>Recopier une formule jusqu'en bas du tableau</td></tr>
  <tr><td><kbd>F4</kbd></td><td>Ajouter ou changer les $ d'une référence</td></tr>
  <tr><td><kbd>F2</kbd></td><td>Modifier la cellule et voir ses références en couleur</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>Maj</kbd> + <kbd>L</kbd></td><td>Activer ou retirer les filtres</td></tr>
  <tr><td><kbd>Ctrl</kbd> + <kbd>!</kbd></td><td>Format nombre avec séparateur (sur certains claviers : Ctrl + Maj + 1)</td></tr>
  <tr><td><kbd>Alt</kbd> + <kbd>Entrée</kbd></td><td>Retour à la ligne dans une cellule</td></tr></tbody></table>
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
