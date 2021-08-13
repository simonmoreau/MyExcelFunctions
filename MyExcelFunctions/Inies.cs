using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelFunctions
{
    class Inies
    {

    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
    public class TINDICATEUR
    {
        public int? ID_INDICATEUR { get; set; }
        public string NOM_INDICATEUR { get; set; }
        public string NOM_COURT { get; set; }
        public int? ID_UNITE { get; set; }
        public int? ID_INDICATEUR_TYPE { get; set; }
        public string CODE_INDICATEUR { get; set; }
        public string NOM_COURT_EN { get; set; }
        public string NOM_INDICATEUR_EN { get; set; }
        public TINDICATEURTYPEs T_INDICATEUR_TYPEs { get; set; }
    }

    public class TINDICATEURs
    {
        public int? ID_INDICATEUR { get; set; }
        public string NOM_INDICATEUR { get; set; }
        public string NOM_COURT { get; set; }
        public int? ID_UNITE { get; set; }
        public int? ID_INDICATEUR_TYPE { get; set; }
        public string CODE_INDICATEUR { get; set; }
        public string NOM_COURT_EN { get; set; }
        public string NOM_INDICATEUR_EN { get; set; }
        public TINDICATEURTYPEs T_INDICATEUR_TYPEs { get; set; }
    }

    public class TINDICATEURTYPEs
    {
        public List<TINDICATEUR> T_INDICATEURs { get; set; }
        public int? ID_INDICATEUR_TYPE { get; set; }
        public string NOM_TYPE { get; set; }
        public string NOM_TYPE_EN { get; set; }
    }

    public class TINDICATEURNORME
    {
        public TINDICATEURs T_INDICATEURs { get; set; }
        public int? ID_INDICATEUR_NORME { get; set; }
        public int? ID_NORME { get; set; }
        public int? ID_INDICATEUR { get; set; }
        public int? ORDRE { get; set; }
        public bool? IS_OPTIONAL { get; set; }
    }

    public class TPHASEs
    {
        public int? ID_PHASE { get; set; }
        public string NOM_PHASE { get; set; }
        public int? TYPE { get; set; }
        public string CODE_PHASE { get; set; }
        public string COMMENT { get; set; }
        public string NOM_PHASE_EN { get; set; }
        public string COMMENT_EN { get; set; }
    }

    public class TPHASENORME
    {
        public TPHASEs T_PHASEs { get; set; }
        public int? ID_PHASE_NORME { get; set; }
        public int? ID_PHASE { get; set; }
        public int? ID_NORME { get; set; }
        public int? ORDRE { get; set; }
        public bool? IS_OPTIONNEL { get; set; }
    }

    public class UNITE
    {
        public List<object> T_CONSTITUANT_UFs { get; set; }
        public List<object> T_INDICATEURs { get; set; }
        public List<object> T_PRODUITs { get; set; }
        public int? ID_UNITE { get; set; }
        public string NOM_UNITE { get; set; }
        public bool? IS_UNITE_UF { get; set; }
        public string DESCRIPTION { get; set; }
    }

    public class NORME
    {
        public List<TINDICATEURNORME> T_INDICATEUR_NORMEs { get; set; }
        public List<TPHASENORME> T_PHASE_NORMEs { get; set; }
        public int? ID_NORME { get; set; }
        public string NOM_NORME { get; set; }
        public DateTime? DATE_AJOUT { get; set; }
        public bool? IS_AVAILABLE { get; set; }
    }



    public class TCONSTITUANTUF
    {
        public object T_UNITEs { get; set; }
        public int? ID_CONSTITUANT_UF { get; set; }
        public int? ID_PRODUIT { get; set; }
        public string NOM_CONSTITUANT { get; set; }
        public double? QUANTITE { get; set; }
        public int? ID_UNITE { get; set; }
        public int? NATURE { get; set; }
    }

    public class TDOCUMENT
    {
        public int? ID_DOCUMENT { get; set; }
        public int? ID_PRODUIT { get; set; }
        public string PATH { get; set; }
        public DateTime? UPLOAD_DATE { get; set; }
        public object KEY_SIGNATURE { get; set; }
        public int? DOC_TYPE { get; set; }
    }

    public class TINDICATEURQUANTITE
    {
        public object T_INDICATEUR_NORMEs { get; set; }
        public object T_PHASE_NORMEs { get; set; }
        public int? ID_INDICATEUR_QUANTITE { get; set; }
        public int? ID_PRODUIT { get; set; }
        public double? QUANTITE { get; set; }
        public int? ID_INDICATEUR_NORME { get; set; }
        public int? ID_PHASE_NORME { get; set; }
    }

    public class TPRODUITDATA
    {
        public int? ID_PRODUIT_DATA { get; set; }
        public int? ID_PRODUIT { get; set; }
        public string LANGUE { get; set; }
        public string REF_COMM { get; set; }
        public string UNITE_FONCTIONNELLE { get; set; }
        public string DOMAINE_APPLICATION { get; set; }
        public object CARACT_NON_CONTENU_DANS_UF { get; set; }
        public object E_COV_FORMALDEHYDE { get; set; }
        public string E_RADIOACTIVES { get; set; }
        public object E_FIBRES_PARTICULES { get; set; }
        public object CROISSANCE_FM { get; set; }
        public object AUTRE_QSEI { get; set; }
        public object INFO_EAU_POTABLE { get; set; }
        public string AUTRES_EAUX { get; set; }
        public string CONFORT_H { get; set; }
        public string CONFORT_A { get; set; }
        public string CONFORT_V { get; set; }
        public string CONFORT_O { get; set; }
        public object AUTRE_CONFORT { get; set; }
        public string DECLARATION_CONTENU { get; set; }
        public object REGLES_EXTRAPOLATION { get; set; }
    }

    public class Produit
    {
        public List<TCONSTITUANTUF> T_CONSTITUANT_UFs { get; set; }
        public List<TDOCUMENT> T_DOCUMENTs { get; set; }
        public List<TINDICATEURQUANTITE> T_INDICATEUR_QUANTITEs { get; set; }
        public object T_NOMENCLATURE_ITEMs { get; set; }
        public object T_OPERATEURs { get; set; }
        public List<TPRODUITDATA> T_PRODUIT_DATAs { get; set; }
        public int? ID_PRODUIT { get; set; }
        public int? ID_ADMIN { get; set; }
        public int? ID_NOMENCLATURE { get; set; }
        public object ID_VALIDATED_FIELDS { get; set; }
        public string NOM_PRODUIT { get; set; }
        public string NATIONAL_KEY { get; set; }
        public DateTime? DATE_VERSION { get; set; }
        public DateTime? DATE_CREATION { get; set; }
        public DateTime? DATE_MILLESIME { get; set; }
        public int? STATUT { get; set; }
        public object MOTS_CLEFS { get; set; }
        public int? MILLESIME { get; set; }
        public int? VERSION { get; set; }
        public object DATE_ARCHIVAGE { get; set; }
        public int? MODIFICATION_LEVEL_TYPE { get; set; }
        public DateTime? DATE_ACV { get; set; }
        public DateTime? DATE_VERIFICATION { get; set; }
        public bool? IS_PUBLIC { get; set; }
        public object ADRESSE_SITE_PUBLIC { get; set; }
        public int? TYPE_GESTION_RATTACHEMENT { get; set; }
        public DateTime? DATE_DEMANDE_VALIDATION { get; set; }
        public DateTime? DATE_VALIDATION { get; set; }
        public int? PRODUIT_TYPE { get; set; }
        public object HAS_LINK_THIRD_PART_TOOL { get; set; }
        public int? DVT { get; set; }
        public double? QUANTITE_UF { get; set; }
        public int? ID_UNIT_UF { get; set; }
        public int? ID_NORME { get; set; }
        public object AUTRE_NORME { get; set; }
        public object NUM_CONFORMITE { get; set; }
        public object IS_NUM_CONF_EXIST { get; set; }
        public bool? IS_CONTACT_EAU_POTABLE { get; set; }
        public bool? IS_CONTACT_EAU_NON_POTABLE { get; set; }
        public double? CHUTE_MEO { get; set; }
        public double? FREQUENCE_ENTRETIEN { get; set; }
        public int? NB_REF_COMM { get; set; }
        public int? NOTATION_AIR_INTERIEUR { get; set; }
        public object IS_PEP { get; set; }
        public object IS_EQUIPEMENT_RT { get; set; }
        public int? VERIFICATION { get; set; }
        public object ID_NOMENCLATURE_2 { get; set; }
        public object ID_NOMENCLATURE_3 { get; set; }
        public object LIEU_PRODUCTION { get; set; }
        public object REGION_PRODUCTION_1 { get; set; }
        public object REGION_PRODUCTION_2 { get; set; }
        public object REGION_PRODUCTION_3 { get; set; }
        public object REGION_PRODUCTION_4 { get; set; }
        public object REGION_PRODUCTION_5 { get; set; }
        public object REGION_PRODUCTION_6 { get; set; }
        public object REGION_PRODUCTION_7 { get; set; }
        public object REGION_PRODUCTION_8 { get; set; }
        public object REGION_PRODUCTION_9 { get; set; }
        public object REGION_PRODUCTION_10 { get; set; }
        public object REGION_PRODUCTION_11 { get; set; }
        public object REGION_PRODUCTION_12 { get; set; }
        public object REGION_PRODUCTION_13 { get; set; }
        public object REGION_PRODUCTION_14 { get; set; }
        public object B_TO_B { get; set; }
        public object IS_DECLARED_PRODUIT { get; set; }
        public object PERF_UF { get; set; }
        public object PERF_UF_DETAILS { get; set; }
        public object ID_UNITE_PERF_UF { get; set; }
        public object QT_PERF_UF { get; set; }
        public object REJECTED_COMMENT { get; set; }
        public object REJECTED_DATE { get; set; }
        public double? DISTANCE_A4 { get; set; }
        public object DISTANCE_C2_DECHETS_RECYCLES { get; set; }
        public object DISTANCE_C2_DECHETS_VALORISES { get; set; }
        public object DISTANCE_C2_DECHETS_ELIMINES { get; set; }
        public object VERIFIED_BY_AFNOR { get; set; }
        public object NUMERO_ENREGISTREMENT { get; set; }
        public object NOM_VERIFICATEUR { get; set; }
        public object BACKGROUND_DATABASE { get; set; }
        public object DATE_ENREGISTREMENT { get; set; }
        public object CARBONE_BIO { get; set; }
        public object IS_FICHE_CONFIGURABLE { get; set; }
        public object IS_FICHE_CONFIGUREE { get; set; }
        public object NUM_ENREGISTREMENT_PARENT { get; set; }
    }


}
