<?php

namespace App\Http\Controllers;

use App\Models\Paiement;
use App\Models\Compte;
use App\Models\Beneficiaire;
use App\Models\Taxe;
use Illuminate\Http\Request;
use App\Exports\PaiementExport;
use App\Imports\PaiementImport;
use Maatwebsite\Excel\Facades\Excel;

use Kwn\NumberToWords\NumberToWords;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

class PaiementController extends Controller
{
    public function index()
    {
        $paiements = Paiement::all();
        return view('paiements.index', compact('paiements'));
    }

    public function create()
    {
        $comptes = Compte::all(); 
        $beneficiaires = Beneficiaire::all(); 
        return view('paiements.create', compact('comptes', 'beneficiaires'));
    }

    public function store(Request $request) 
    {
        // $request->validate([
        //     'montant' => 'required|numeric',
        //     'date_paiement' => 'required|date',
        //     'mode_paiement' => 'required|in:carte,virement,cheque,espèces',
        //     'id_compte' => 'required|exists:compte,id',
        //     'id_beneficiaire' => 'required|exists:beneficiaire,id',
        //     'status' => 'required|in:en attente,réussi,échoué',
        //     'motif_de_la_depence' => 'required|text',
        //     'impulsion' => 'required|in:TVA,IMF,loyer,Exonéré',
        // ]);
    
        // Paiement::create([
        //     'montant' => $request->montant,
        //     'date_paiement' => $request->date_paiement,
        //     'mode_paiement' => $request->mode_paiement,
        //     'id_compte' => $request->id_compte,
        //     'id_beneficiaire' => $request->id_,  
        //     'status' => $request->status,
        // ]);
    
        return redirect()->route('paiements.index')->with('success', 'Paiement ajouté avec succès!');
    }
    

    

    public function edit($id)
    {
        $paiement = Paiement::findOrFail($id);
        $comptes = Compte::all();
        $beneficiaires = Beneficiaire::all();
        return view('paiements.edit', compact('paiement', 'comptes', 'beneficiaires'));
    }

    public function update(Request $request, Paiement $paiement)
    {
        // Validation des données
        $request->validate([
            'montant' => 'required|numeric',
            'mode_paiement' => 'required|in:carte,virement,cheque,espèces',
            'id_compte' => 'required|exists:compte,id',
            'id_beneficiaire' => 'required|exists:beneficiaire,id',
            'status' => 'required|in:en attente,réussi,échoué',
            'motif_de_la_depence' => 'required|string',
            'impulsion' => 'required|in:TVA,IMF,loyer,Exonéré',
        ]);

        $paiement->update($request->all());
        // Rediriger avec un message de succès
        return redirect()->route('paiements.index')->with('success', 'Paiement mis à jour avec succès.');
    }


    public function destroy($id)
    {
        $paiement = Paiement::findOrFail($id);
        $paiement->delete();

        return redirect()->route('paiements.index')->with('success', 'Paiement supprimé avec succès.');
    }

    public function export($id)
    {
        $paiement = Paiement::findOrFail($id);
        return Excel::download(new PaiementExport, 'paiements.xlsx');
    }

    public function import(Request $request)
    {
    $request->validate([
        'file' => 'required|mimes:xlsx,csv',
    ]);

    Excel::import(new PaiementImport, $request->file('file'));

    return back()->with('success', 'Les paiements ont été importés avec succès.');
    }

    function convertirMontantEnLettres($montant) {
        $montant = number_format($montant, 2, '.', ''); // Assurez-vous que le montant est bien formaté avec deux décimales
    
        $chiffres = [
            0 => 'zéro', 1 => 'un', 2 => 'deux', 3 => 'trois', 4 => 'quatre', 5 => 'cinq', 6 => 'six', 
            7 => 'sept', 8 => 'huit', 9 => 'neuf', 10 => 'dix', 11 => 'onze', 12 => 'douze', 13 => 'treize', 
            14 => 'quatorze', 15 => 'quinze', 16 => 'seize', 17 => 'dix-sept', 18 => 'dix-huit', 19 => 'dix-neuf', 
            20 => 'vingt', 30 => 'trente', 40 => 'quarante', 50 => 'cinquante', 60 => 'soixante', 
            70 => 'soixante-dix', 80 => 'quatre-vingts', 90 => 'quatre-vingt-dix'
        ];
    
        // Conversion de la partie entière
        $entier = floor($montant);
        $decimales = round(($montant - $entier) * 100); // On prend les deux premières décimales
    
        // Conversion des centaines, milliers, etc.
        if ($entier == 0) {
            return $chiffres[0];
        }
    
        $lettres = '';
    
        // Des milliers et centaines
        if ($entier >= 1000) {
            $milliers = floor($entier / 1000);
            $lettres .= $this->convertirMontantEnLettres($milliers) . ' mille ';
            $entier %= 1000;
        }
    
        if ($entier >= 100) {
            $centaines = floor($entier / 100);
            if ($centaines > 1) {
                $lettres .= $chiffres[$centaines] . ' cent ';
            } else {
                $lettres .= 'cent ';
            }
            $entier %= 100;
        }
    
        // Les dizaines et unités
        if ($entier >= 20) {
            $dizaines = floor($entier / 10) * 10;
            $lettres .= $chiffres[$dizaines] . ' ';
            $entier %= 10;
        }
    
        if ($entier > 0) {
            $lettres .= $chiffres[$entier];
        }
    
        // Conversion de la partie décimale (centimes)
        if ($decimales > 0) {
            $lettres .= ' et ' . $decimales . '/100';
        }
    
        return ucfirst(trim($lettres)); // Capitalize the first letter for better formatting
    }
    
    public function show($id)
    {
        $comptes = Compte::all(); 
        $beneficiaires = Beneficiaire::all();
        $paiement = Paiement::findOrFail($id);
        $taxes = Taxe::all();
        
        $tva = $imf = $pl = $cf = $irf = 0;

        foreach ($taxes as $taxe) {
            if ($taxe->nom == 'TVA') {
                $tva = $taxe->pourcentage / 100;
            } elseif ($taxe->nom == 'IMF') {
                $imf = $taxe->pourcentage / 100;
            } elseif ($taxe->nom == 'PL') {
                $pl = $taxe->pourcentage / 100;
            } elseif ($taxe->nom == 'CF') {
                $cf = $taxe->pourcentage / 100;
            } elseif ($taxe->nom == 'IRF') {
                $irf = $taxe->pourcentage / 100;
            }
        }

        if($paiement->impulsion == 'TVA'){

            $phpWord = new PhpWord();
            $section = $phpWord->addSection();
        
            // Add content to the Word document
            $section->addText("République Islamique de Mauritanie:", null, ['align' => Jc::LEFT]);
            $section->addText("Honneur-Fraternité-Justice:", null, ['align' => Jc::LEFT]);
            $section->addText("MINISTERE DE L'ENSEIGNEMENT SUPERIEUR:", null, ['align' => Jc::LEFT]);
            $section->addText("ET DE LA RECHERCHE SCIENTIFIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("INSTITUT SUPERIEUR NUMERIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("Titre de paiement Numero:" , null, ['align' => Jc::CENTER]);

            $section->addText("Imputation budgetaire: Compte principale ");
            $section->addText("Bénéficiaire: "); 
            $section->addText("Montant: " . $paiement->montant);
            $section->addText("Montant en lettres: " . $this->convertirMontantEnLettres($paiement->montant));

            $section->addText("Mode de Paiement: " . $paiement->mode_paiement);
            
            $TCC = $paiement->montant;
            $HT = $TCC/(1 + $tva);
            $calc_tva = ($tva * $TCC)/(1 + $tva);
            $calc_imf = $TCC * $imf ;
            $net = $TCC - ($tva + $imf);

            $section->addText("TVA: " .$calc_tva);
            $section->addText("IMF: " .$calc_imf);
            $section->addText("ITS: ");
            $section->addText("CNAM: ");
            $section->addText("NET: " .$net);
            
            $section->addText("La date: " . $paiement->date_paiement, null, ['align' => Jc::RIGHT]);
            $section->addText("Le Directeur", ['align' => Jc::RIGHT]);
            $section->addText("Le Comptable", ['align' => Jc::LEFT]);
            
            $fileName = 'paiement_' . $paiement->id . '.docx';
            $filePath = storage_path('app/public/' . $fileName);
            $phpWord->save($filePath, 'Word2007');
        
            // Envoyer le fichier Word pour l'ouvrir dans le navigateur
            return response()->file($filePath, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition' => 'inline; filename="' . $fileName . '"'
            ])->deleteFileAfterSend(true);  // Suppression du fichier après l'envoi

        }
        elseif($paiement->impulsion == 'IMF'){

            $phpWord = new PhpWord();
            $section = $phpWord->addSection();
        
            // Add content to the Word document
            $section->addText("République Islamique de Mauritanie:", null, ['align' => Jc::LEFT]);
            $section->addText("Honneur-Fraternité-Justice:", null, ['align' => Jc::LEFT]);
            $section->addText("MINISTERE DE L'ENSEIGNEMENT SUPERIEUR:", null, ['align' => Jc::LEFT]);
            $section->addText("ET DE LA RECHERCHE SCIENTIFIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("INSTITUT SUPERIEUR NUMERIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("Titre de paiement Numero:" , null, ['align' => Jc::CENTER]);

            $section->addText("Imputation budgetaire: Compte principale ");
            $section->addText("Bénéficiaire: ");
            $section->addText("Montant: " . $paiement->montant);
            $section->addText("Montant en lettres: " . $this->convertirMontantEnLettres($paiement->montant));

            $section->addText("Mode de Paiement: " . $paiement->mode_paiement);

            $TTC = $paiement->montant;
            $calc_imf = $TTC * $imf;
            $net = $TTC - $calc_imf;

            $section->addText("TVA: ");
            $section->addText("IMF: " .$calc_imf);
            $section->addText("ITS: ");
            $section->addText("CNAM: ");
            $section->addText("NET: " .$net);
            
            $section->addText("La date: " . $paiement->date_paiement);
            $section->addText("Le Directeur", ['align' => Jc::RIGHT]);
            $section->addText("Le Comptable", ['align' => Jc::LEFT]);

            $fileName = 'paiement_' . $paiement->id . '.docx';
            $filePath = storage_path('app/public/' . $fileName);
            $phpWord->save($filePath, 'Word2007');

            // Envoyer le fichier Word pour l'ouvrir dans le navigateur
            return response()->file($filePath, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition' => 'inline; filename="' . $fileName . '"'
            ])->deleteFileAfterSend(true);  // Suppression du fichier après l'envoi

        }

        elseif($paiement->impulsion == 'Loyer'){
            $phpWord = new PhpWord();
            $section = $phpWord->addSection();
        
            // Add content to the Word document
            $section->addText("République Islamique de Mauritanie:", null, ['align' => Jc::LEFT]);
            $section->addText("Honneur-Fraternité-Justice:", null, ['align' => Jc::LEFT]);
            $section->addText("MINISTERE DE L'ENSEIGNEMENT SUPERIEUR:", null, ['align' => Jc::LEFT]);
            $section->addText("ET DE LA RECHERCHE SCIENTIFIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("INSTITUT SUPERIEUR NUMERIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("Titre de paiement Numero:" , null, ['align' => Jc::CENTER]);

            $section->addText("Imputation budgetaire: Compte principale ");
            $section->addText("Bénéficiaire: ");
            $section->addText("Montant: " . $paiement->montant);
            $section->addText("Montant en lettres: " . $this->convertirMontantEnLettres($paiement->montant));

            $section->addText("Mode de Paiement: " . $paiement->mode_paiement);
            
            

            $calc_pl = ($paiement->montant) * $pl;
            $calc_cf = ($paiement->montant) * $cf;
            $calc_irf = ($paiement->montant) * $irf;

            $net = ($paiement->montant) - ($calc_pl + $calc_cf + $calc_irf);

            $section->addText("TVA: " );
            $section->addText("PL: " .$calc_pl);
            $section->addText("CF: " .$calc_cf);
            $section->addText("IRF: " .$calc_irf);
            $section->addText("NET: " .$net);
            $section->addText("La date: " . $paiement->date_paiement);
            $section->addText("Le Directeur", ['align' => Jc::RIGHT]);
            $section->addText("Le Comptable", ['align' => Jc::LEFT]);

            // Save the Word document
            $fileName = 'paiement_' . $paiement->id . '.docx';
            $filePath = storage_path('app/public/' . $fileName);
            $phpWord->save($filePath, 'Word2007');

            // Envoyer le fichier Word pour l'ouvrir dans le navigateur
            return response()->file($filePath, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition' => 'inline; filename="' . $fileName . '"'
            ])->deleteFileAfterSend(true);  // Suppression du fichier après l'envoi
        }

        else{
            $phpWord = new PhpWord();
            $section = $phpWord->addSection();
            
                // Add content to the Word document
            $section->addText("République Islamique de Mauritanie:", null, ['align' => Jc::LEFT]);
            $section->addText("Honneur-Fraternité-Justice:", null, ['align' => Jc::LEFT]);
            $section->addText("MINISTERE DE L'ENSEIGNEMENT SUPERIEUR:", null, ['align' => Jc::LEFT]);
            $section->addText("ET DE LA RECHERCHE SCIENTIFIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("INSTITUT SUPERIEUR NUMERIQUE:" , null, ['align' => Jc::LEFT]);
            $section->addText("Titre de paiement Numero:" , null, ['align' => Jc::CENTER]);
    
            $section->addText("Imputation budgetaire: Compte principale ");
            $section->addText("Bénéficiaire: ");
            $section->addText("Montant: " . $paiement->montant);
            $section->addText("Montant en lettres: " . $this->convertirMontantEnLettres($paiement->montant));
    
            $section->addText("Mode de Paiement: " . $paiement->mode_paiement);
            
            $section->addText("TVA: ");
            $section->addText("PL: ");
            $section->addText("CF: ");
            $section->addText("IRF: ");
            $section->addText("NET: " .$paiement->montant);

            
                
            $section->addText("La date: " . $paiement->date_paiement);
            $section->addText("Le Directeur", ['align' => Jc::RIGHT]);
            $section->addText("Le Comptable", ['align' => Jc::LEFT]);

            $fileName = 'paiement_' . $paiement->id . '.docx';
            $filePath = storage_path('app/public/' . $fileName);
            $phpWord->save($filePath, 'Word2007');

            // Envoyer le fichier Word pour l'ouvrir dans le navigateur
            return response()->file($filePath, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition' => 'inline; filename="' . $fileName . '"'
            ])->deleteFileAfterSend(true);  // Suppression du fichier après l'envoi

            
        }

    }

    
}
