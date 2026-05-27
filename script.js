/* =========================================================
           APP NAMESPACE — v2 (no locking, role-based, real-time sync)
           All state, master data, persistence, and UI bound here.
        ========================================================= */
const App = (function () {
  'use strict';

  /* =========================================================
               1. CONFIGURATION
               =========================================================
               To enable live Pan-India sync between admins and IC members,
               set FIREBASE_URL below to your Firebase Realtime Database URL.
               Setup steps (one-time, by admin):
                 1. Create a free Firebase project (https://console.firebase.google.com)
                 2. Enable "Realtime Database" → start in test mode for trial
                 3. Copy the database URL (looks like https://<project>-default-rtdb.firebaseio.com)
                 4. Paste it below as FIREBASE_URL
                 5. Save the file and redeploy — sync will activate automatically
               If FIREBASE_URL is left empty, the app falls back to local storage
               plus same-device BroadcastChannel sync between browser tabs, with
               manual JSON export/import for cross-device data movement.
            ========================================================= */
  const KEY_FB_URL = 'axis_burgundy_ic_firebase_url_v3';

  /* Resolve the Firebase Realtime DB URL from (in order):
                 1. ?fb=<url> URL parameter (shareable pre-config)
                 2. localStorage (set once via the Master page)
                 3. Empty → falls back to local-only mode
               Priority lets an admin share a link like
                 index.html?fb=https://my-rtdb.firebaseio.com
               so all recipients are auto-configured on first visit. */
  function resolveFirebaseURL() {
    try {
      const qs = new URLSearchParams(window.location.search);
      const fromQuery = (qs.get('fb') || '').trim();
      if (fromQuery) {
        localStorage.setItem(KEY_FB_URL, fromQuery);
        return fromQuery;
      }
    } catch (e) {}
    try {
      return (localStorage.getItem(KEY_FB_URL) || '').trim();
    } catch (e) {
      return '';
    }
  }

  const CONFIG = {
    FIREBASE_URL:
      'https://investment-tracker-d4d88-default-rtdb.asia-southeast1.firebasedatabase.app', // ← set once via Master page, or via ?fb= URL param
    CHANNEL: 'axis-burgundy-ic-sync-v2', // BroadcastChannel name
    POLL_MS: 12000, // Fallback polling interval if SSE unavailable
    // === AUTHORISED ADMINS (per business spec, locked list) ===
    // Only these four individuals can hold the ADMIN role.
    // Aliases (alternate spellings used in the Master Excel) are accepted
    // by the authorisation check so admins are recognised whichever form
    // their name appears under in HR / Master records.
    ADMIN_ALLOWLIST: [
      'Kaukab Shamshi',
      'Tarun Dhingra',
      'Rajat Mohapatra',
      'Chethan C',
    ],
    ADMIN_ALIASES: {
      // canonical name → list of acceptable spellings
      'Kaukab Shamshi': ['Kaukab Shamsi'],
      'Tarun Dhingra': ['Tarun W'],
      'Rajat Mohapatra': [
        'Rajat NE',
        'Rajat Kumar Mohapatra',
        'RajatK Mohapatra',
      ],
      'Chethan C': ['Chethan S'],
    },
  };

  /* Persist a new Firebase URL and re-initialise sync without a page reload */
  function setFirebaseURL(url) {
    url = (url || '').trim().replace(/\/+$/, '');
    if (url && !/^https?:\/\//i.test(url)) {
      toast('error', 'Invalid URL', 'Firebase URL must start with https://');
      return false;
    }
    try {
      localStorage.setItem(KEY_FB_URL, url);
    } catch (e) {}
    CONFIG.FIREBASE_URL = url;
    logAudit(
      'SYNC_CONFIG',
      '—',
      url ? 'Firebase URL configured' : 'Firebase URL cleared',
    );
    // Re-initialise sync layer to pick up the new URL
    try {
      Sync.init();
    } catch (e) {}
    toast(
      'info',
      url ? 'Sync enabled' : 'Sync disabled',
      url
        ? 'Real-time sync is now active across all admins.'
        : 'Reverted to local-only mode.',
    );
    return true;
  }

  /* =========================================================
               2. DEFAULT MASTER DATA
               Derived from Master_RM_IC_Data Excel (RM_Data + IC_Zone_Map).
               Replaced at runtime when Admin uploads a fresh Master file.
            ========================================================= */
  // const DEFAULT_MASTER = {
  //   /* ==========================================================
  //                  DATA BELOW IS AUTO-GENERATED FROM Master_RM_IC_Data.xlsx
  //                  To refresh: re-run build.py with an updated Excel.
  //                  ICs: 23 | RICs: 4 | Partners: 159
  //               ========================================================== */

  //   /* Exactly three zones per business policy */
  //   zones: ['South Zone', 'West Zone', 'North-East Region'],

  //   /* 5 Meeting Types — unchanged from v1 */
  //   meetingTypes: [
  //     'Portfolio Review',
  //     'Investment Pitch',
  //     'Handover',
  //     'Documentation',
  //     'Others',
  //   ],

  //   /* Three Client Types (Burgundy segmentation) */
  //   clientTypes: ['Existing to Bank', 'New to Bank', 'New to Investment'],

  //   /* Exactly 5 Product Categories (RESTRICTED — no other values permitted) */
  //   categories: [
  //     'AIF/PMS',
  //     'Mutual Funds',
  //     'Private Equity',
  //     'Multiple Products',
  //     'Others',
  //   ],

  //   /* RIC heads (from Master Dropdowns_Ref + RM_Data) */
  //   ricMembers: [
  //     'Chethan C',
  //     'Kaukab Shamshi',
  //     'Rajat Kumar Mohapatra',
  //     'Tarun Dhingra',
  //   ],

  //   /* IC → RIC mapping (from Master RIC_NAME column) */
  //   icToRic: {
  //     'Aditya Samant Kapoor': 'Tarun Dhingra',
  //     'Amit Khanna': 'Tarun Dhingra',
  //     'Anupam Lal': 'Tarun Dhingra',
  //     'Anurag Minj': 'Kaukab Shamshi',
  //     'Bhaskar Rautela': 'Rajat Kumar Mohapatra',
  //     'Chethan C': 'Kaukab Shamshi',
  //     'Deepak Narain Chellani': 'Rajat Kumar Mohapatra',
  //     'Hedley Albuquerque': 'Kaukab Shamshi',
  //     'Juvin I C': 'Kaukab Shamshi',
  //     'Mallika Swarup': 'Tarun Dhingra',
  //     'Mayank Mehta': 'Tarun Dhingra',
  //     'Mudholkar Shrirang D': 'Rajat Kumar Mohapatra',
  //     'Nirmalya Majumdar': 'Tarun Dhingra',
  //     'Prince Pathania': 'Tarun Dhingra',
  //     'Rajat Kumar Mohapatra': 'Rajat Kumar Mohapatra',
  //     'Ranjeet Kumar Jha': 'Tarun Dhingra',
  //     'Rupesh Kumar': 'Kaukab Shamshi',
  //     'Shashwat Gupta': 'Tarun Dhingra',
  //     'Suvid Nazar': 'Rajat Kumar Mohapatra',
  //     'Tarun Dhingra': 'Tarun Dhingra',
  //     'Vijay Karumanchi': 'Kaukab Shamshi',
  //     'Vikalp Kamlesh Shah': 'Kaukab Shamshi',
  //     'Vishal Balsara': 'Rajat Kumar Mohapatra',
  //   },

  //   /* IC → Zone mapping (bucketed to 3 zones from Master Region + IC_Zone_Map) */
  //   icToZone: {
  //     'Aditya Samant Kapoor': 'North-East Region',
  //     'Amit Khanna': 'North-East Region',
  //     'Anupam Lal': 'North-East Region',
  //     'Anurag Minj': 'South Zone',
  //     'Bhaskar Rautela': 'West Zone',
  //     'Chethan C': 'South Zone',
  //     'Deepak Narain Chellani': 'West Zone',
  //     'Hedley Albuquerque': 'West Zone',
  //     'Juvin I C': 'South Zone',
  //     'Mallika Swarup': 'North-East Region',
  //     'Mayank Mehta': 'North-East Region',
  //     'Mudholkar Shrirang D': 'West Zone',
  //     'Nirmalya Majumdar': 'North-East Region',
  //     'Prince Pathania': 'North-East Region',
  //     'Rajat Kumar Mohapatra': 'West Zone',
  //     'Ranjeet Kumar Jha': 'North-East Region',
  //     'Rupesh Kumar': 'South Zone',
  //     'Shashwat Gupta': 'North-East Region',
  //     'Suvid Nazar': 'West Zone',
  //     'Tarun Dhingra': 'North-East Region',
  //     'Vijay Karumanchi': 'South Zone',
  //     'Vikalp Kamlesh Shah': 'West Zone',
  //     'Vishal Balsara': 'West Zone',
  //   },

  //   /* IC → [Partner …] mapping (from Master RM_NAME column) */
  //   icToPartners: {
  //     'Aditya Samant Kapoor': [
  //       'Roopsi Anand',
  //       'Neeraj Krishan Rajput',
  //       'Nishchay Jagdish Jethanandani',
  //       'Mayank Mehta',
  //       'Tarun Agarwal',
  //       'Nandita Kulbhushan Vats',
  //       'Pranjal Dwivedi',
  //       'Raghav Khanna',
  //       'Sachin Bhardwaj',
  //     ],

  //     'Amit Khanna': [
  //       'Meenakshi Rudresh J',
  //       'Azimul Haque',
  //       'Bharat Popli',
  //       'Sonali Dhingra',
  //       'Namrata Madan Vidyarthi',
  //       'Harsh Monga',
  //       'Nilesh Sarkar',
  //       'Akshansh Joshi',
  //     ],

  //     'Anupam Lal': [
  //       'Varun Jawa',
  //       'Purnima Jain',
  //       'Priya Vig',
  //       'Ravi Mohan Srivastava',
  //       'Manuj Suri',
  //       'Lalit Goyal',
  //       'Avnish Joshi',
  //     ],

  //     'Anurag Minj': [
  //       'Tirupathi Marella',
  //       'Santosh Kumar A',
  //       'Sukarno Jami',
  //       'Rajeev Tripathy',
  //       'Revanth Kumar Matukumalli',
  //       'N Sudhakar',
  //       'Aysha Amreen',
  //       'Venkataraman Natarajan',
  //       'Rahulrao Ramanarao Panduru',
  //       'P Sreekanth',
  //       'Durga Hari',
  //       'Sandeep Reddy Vontikommu',
  //       'Vasudebo Naram',
  //       'Kiran Kumar Alli',
  //       'Chandu Gopi Krishna Kanth',
  //       'Venkata Satheesh Chitta',
  //       'Sindhu K M Shet',
  //       'John Rohan Sequeira',
  //       'Nandini Reddy Bhimavarapu',
  //       'Shalini B V',
  //     ],

  //     'Bhaskar Rautela': [
  //       'Anish M Desai',
  //       'Ninad Narayan Salunkhe',
  //       'Abhishek Vinod Beri',
  //       'Taizun Mohammed Badshah',
  //       'Yogesh Gupta',
  //       'Richa Dhasmana',
  //       'Deepak Suryawanshi',
  //       'Shilpa Chadha',
  //       'Akshay Balasaheb Patil',
  //       'Nikhil Karn Puranik',
  //     ],

  //     'Chethan C': [
  //       'Thrishul Kanthreddy',
  //       'C K Raghu',
  //       'Samreen Nawab',
  //       'Senthil Kasiviswanathan',
  //       'Kavitha Rai',
  //       'Phillips Kurian',
  //       'Ankit Choudhary',
  //       'Prabhakara S',
  //       'Ramesh Rajasekar',
  //     ],

  //     'Deepak Narain Chellani': [
  //       'Iqbal Mehdi',
  //       'Ambrish Poonamchand Asopa',
  //       'Jagdish K Biyani',
  //       'Vikas P Jadhav',
  //       'Sachin Kutty Shetty',
  //       'Mahendra Parihar',
  //       'Dinesh Arora',
  //       'Ankur Sanjay Chemburkar',
  //       'Chaitanya Pramod Koli',
  //       'Viraj Nishith Panchmia',
  //       'Sandeep Kalra',
  //       'Sreejita Goswami',
  //       'Mahak Manoj Khanna',
  //     ],

  //     'Hedley Albuquerque': [
  //       'Farzoon Qureshi',
  //       'Kaushlendra Makhu Singh',
  //       'Dinesh Gobind Jhangiani',
  //       'Rajesh Ramesh Chaturvedi',
  //       'Mayuresh Narottam Bharadwa',
  //       'Amit Kumar Singh',
  //       'Sonia Verma',
  //       'Ajay Kaushal',
  //       'Ketan Amrut Vora',
  //       'Shivani Dalal',
  //       'Devesh Srivastava',
  //       'Kamal Hasant Shroff',
  //       'Shanti Vijay Shukla',
  //     ],

  //     'Juvin I C': [
  //       'Karthic M',
  //       'K Dinesh',
  //       'Arjun Bhaskar',
  //       'Sreeram .H',
  //       'Divya Sajil',
  //       'Nithin Rajan Subalaja',
  //       'Jibi M Mathew Jacob',
  //       'Karthikeyan G',
  //       'Antlin Jenifer R',
  //       'Vijikumar V C',
  //     ],

  //     'Mallika Swarup': [
  //       'Gurpreet Singh',
  //       'Sandeep Kaura',
  //       'Kunal Bahl',
  //       'Priyanka Sharma',
  //       'Neeraj Kachru',
  //       'Rohit Sinha',
  //       'Parul Sharma',
  //       'Seemant Sharma',
  //     ],

  //     'Mayank Mehta': [
  //       'Amit Sharma',
  //       'Ankur Garg',
  //       'Syed Bilal Hussain',
  //       'Nidhi Jaiswal',
  //       'Kashish Kapoor',
  //       'Akanksha Narain',
  //     ],

  //     'Mudholkar Shrirang D': [
  //       'Dhananjay Nalinkant Mehta',
  //       'Niyati Majithia',
  //       'Raj Kumar Jaiswal',
  //       'Sachin Kirtikumar Vora',
  //       'Shreepad Uddhavrao Wakodkar',
  //       'Rashi Naman Shukla',
  //       'Umed Singh Rao',
  //       'Abhani Shreyas Satishbhai',
  //       'Jigar P Pathak',
  //       'Raithatha Amit',
  //       'Anshul Jain',
  //       'Muskaan Madan',
  //     ],

  //     'Nirmalya Majumdar': [
  //       'Soumyajit Ghosh',
  //       'Sumona Sengupta Dutta',
  //       'Gourav Doshi',
  //       'Kaunak Biswas',
  //       'Ashish Goenka',
  //       'Kaushik Chakrabarty',
  //       'Somenath Dey',
  //       'Rakhi Kar',
  //     ],

  //     'Prince Pathania': [
  //       'Dasmeet Singh Kohli',
  //       'Mansi Mahajan',
  //       'Ajay Bakshi',
  //       'Kanuj Bharti',
  //       'Akash Kumar',
  //       'Saurabh Gupta',
  //       'Bhaskar Rajvanshi',
  //       'Anu Ahluwalia',
  //     ],

  //     'Rajat Kumar Mohapatra': ['Pradip Vitthal Jagtap', 'Puneet Panchaity'],

  //     'Ranjeet Kumar Jha': [
  //       'Ankit Jain',
  //       'Ankur Agarwal',
  //       'Abhishek Verma',
  //       'Gaurav Sharma',
  //       'Gaurav Agrahari',
  //       'Indra Kumar Jain',
  //       'Ankit Kanungo',
  //       'Vipin Gupta',
  //       'Dilip Paliwal',
  //       'Kshitij Tiwari',
  //       'Saif Aijaz',
  //       'Amit Audichya',
  //       'Mansi Gambhir',
  //     ],

  //     'Rupesh Kumar': [
  //       'Gurumurthy S',
  //       'Abhishek Mallick',
  //       'Ashish Deepak Mehta',
  //       'Argha Banerjee',
  //       'Soumya Mukherjee',
  //       'K Jyoti',
  //       'Nitesh Dave',
  //       'Harpreet Rohit Nagrale',
  //       'Niharika Azad',
  //       'Akshay Tyagi',
  //       'Himanshu Goyal',
  //       'Sathish T',
  //       'Suruchi Rai',
  //       'Sandeep Kumar Pandey',
  //       'Vivek Yadav',
  //       'Ashish Singh',
  //     ],

  //     'Shashwat Gupta': [
  //       'Manoj Kumar Sarma',
  //       'Sukamal Majee',
  //       'Biswajeet Sengupta',
  //       'Vikash Kedia',
  //       'Siddhant Nawal',
  //       'Debashis Mallick',
  //     ],

  //     'Suvid Nazar': [
  //       'Priyank Amrutbhai Patel',
  //       'Trivedi Tarak Maheshbhai',
  //       'Jaykumar Solanki',
  //       'Divyanand Tiwari',
  //       'Zarna Harshad Zaveri',
  //       'Shahnaz Zubin Mehta',
  //       'Tejas Parekh',
  //       'Jeetendra Kumar',
  //       'Hetal Vinod Shah',
  //       'Chirag Dahyabhai Katariya',
  //       'Himani Hiren Pandya',
  //       'Akash Pareek',
  //       'Nishant Gordhan Thanki',
  //       'Bharat Dilip Kulkarni',
  //     ],

  //     'Tarun Dhingra': [
  //       'Ashish Jain',
  //       'Privanka Sharma',
  //       'Swati Gupta',
  //       'Kapil Chadha',
  //     ],

  //     'Vijay Karumanchi': [
  //       'Suman Chakravarthy S',
  //       'J Santhanalakshmi',
  //       'M Ganesh Kumar',
  //       'Vickram S',
  //       'Kamalnath S',
  //       'Alok Banthia',
  //       'D Vivekanandan',
  //       'Manivannan S',
  //       'C Arun',
  //       'Adyith Santhanam',
  //       'Ruthpriya J',
  //       'Sridhar A',
  //     ],

  //     'Vikalp Kamlesh Shah': [
  //       'Pranot Prakash Patil',
  //       'Sarish Menon',
  //       'Suresh Mogaveera',
  //       'Parijat Ghosh',
  //       'Omprakash Bechan Yadav',
  //       'Shivani Kamal Rawat',
  //       'Jigna Chirag Sheth',
  //       'Ayan Kumar Bhowmick',
  //       'Ashwini Kumar Chanda',
  //       'Lakshmi Narayan',
  //       'Hemal Kanaiyalal Oza',
  //     ],

  //     'Vishal Balsara': [
  //       'Manish Ramesh Balwani',
  //       'Sumit Bhide',
  //       'Pawan Jakhetia',
  //       'Lokesh Singh Chauhan',
  //       'Shantanu Shrivastava',
  //       'Antonio Vicente Ethelbert Dias',
  //       'Shobhit Agrawal',
  //       'Harsh Vardhan',
  //       'Mithila Pawar',
  //     ],
  //   },

  //   /* Full sorted list of IC Members (drives dropdowns) */
  //   icMembers: [
  //     'Aditya Samant Kapoor',
  //     'Amit Khanna',
  //     'Anupam Lal',
  //     'Anurag Minj',
  //     'Bhaskar Rautela',
  //     'Chethan C',
  //     'Deepak Narain Chellani',
  //     'Hedley Albuquerque',
  //     'Juvin I C',
  //     'Mallika Swarup',
  //     'Mayank Mehta',
  //     'Mudholkar Shrirang D',
  //     'Nirmalya Majumdar',
  //     'Prince Pathania',
  //     'Rajat Kumar Mohapatra',
  //     'Ranjeet Kumar Jha',
  //     'Rupesh Kumar',
  //     'Shashwat Gupta',
  //     'Suvid Nazar',
  //     'Tarun Dhingra',
  //     'Vijay Karumanchi',
  //     'Vikalp Kamlesh Shah',
  //     'Vishal Balsara',
  //   ],
  // };

  const GOOGLE_SHEET_ID = '1q_WALEeOXyL-cqAxGtAkxhU4XbifjCJwxTuDOQh09aY';
  const GOOGLE_SHEET_NAME = 'Sheet1';
  // https://opensheet.elk.sh/1q_WALEeOXyL-cqAxGtAkxhU4XbifjCJwxTuDOQh09aY/Sheet1
  // https://docs.google.com/spreadsheets/d/1q_WALEeOXyL-cqAxGtAkxhU4XbifjCJwxTuDOQh09aY/edit?usp=sharing

  const DEFAULT_MASTER = {
    zones: ['South Zone', 'West Zone', 'North-East Region'],

    meetingTypes: [
      'Portfolio Review',
      'Investment Pitch',
      'Handover',
      'Documentation',
      'Winback',
      'Others',
    ],

    clientTypes: ['Existing', 'NTB', 'NTI'],

    categories: [
      'AIF/PMS',
      'Mutual Funds',
      'Unlisted',
      'Multiple Products',
      'Others',
    ],

    ricMembers: [],
    icMembers: [],
    icToRic: {},
    icToZone: {},
    icToPartners: {},
    icPasswords: {},
  };

  function normalizeZone(region) {
    const r = String(region || '')
      .trim()
      .toLowerCase();

    if (r.includes('south')) {
      return 'South Zone';
    }

    if (r.includes('west')) {
      return 'West Zone';
    }

    if (r.includes('north') || r.includes('east') || r.includes('ne')) {
      return 'North-East Region';
    }

    return '';
  }

  async function loadMasterFromGoogleSheet() {
    try {
      const url = `https://opensheet.elk.sh/${GOOGLE_SHEET_ID}/${GOOGLE_SHEET_NAME}`;

      const response = await fetch(url);
      const rows = await response.json();

      console.log('Google Sheet rows:', rows);

      const master = {
        ...DEFAULT_MASTER,
        ricMembers: [],
        icMembers: [],
        icToRic: {},
        icToZone: {},
        icToPartners: {},
      };

      const ricSet = new Set();
      const icSet = new Set();

      rows.forEach((row) => {
        const region = String(
          row.Region || row.REGION || row.region || '',
        ).trim();

        const rmName = String(
          row.RM_NAME || row['RM NAME'] || row.rm_name || '',
        ).trim();

        const icName = String(
          row.IC_NAME || row['IC NAME'] || row.ic_name || '',
        ).trim();

        const ricName = String(
          row.RIC_NAME || row['RIC NAME'] || row.ric_name || '',
        ).trim();

        const password = String(row.password || row.PASSWORD || '').trim();

        if (!icName) return;

        icSet.add(icName);

        if (ricName) {
          ricSet.add(ricName);
          master.icToRic[icName] = ricName;
        }

        if (password) {
          master.icPasswords[icName] = password;
        }

        const normalizedZone = normalizeZone(region);
        if (normalizedZone) {
          master.icToZone[icName] = normalizedZone;
        }

        if (!master.icToPartners[icName]) {
          master.icToPartners[icName] = [];
        }

        if (rmName && !master.icToPartners[icName].includes(rmName)) {
          master.icToPartners[icName].push(rmName);
        }
      });

      master.icMembers = [...icSet].sort();
      master.ricMembers = [...ricSet].sort();

      console.log('IC COUNT:', master.icMembers.length);
      console.log(master.icMembers);

      MASTER = master;

      saveJSON(KEY_MASTER, master);

      populateMaster();

      refreshUserSelect();

      refreshAll();

      toast(
        'info',
        'Google Sheet Loaded',
        `${master.icMembers.length} IC members loaded`,
      );

      console.log('Master data loaded successfully');
    } catch (error) {
      console.error('Failed to load Google Sheet:', error);

      toast('error', 'Google Sheet Error', 'Could not load master data');
    }
  }

  /* Live MASTER reference — may be replaced by uploaded master at runtime */
  let MASTER = null;

  /* =========================================================
               3. STORAGE KEYS & STATE
            ========================================================= */
  const KEY_RECORDS = 'axis_burgundy_ic_records_v2';
  const KEY_AUDIT = 'axis_burgundy_ic_audit_v2';
  const KEY_DRAFT = 'axis_burgundy_ic_draft_v2';
  const KEY_SEQ = 'axis_burgundy_ic_seq_v2'; // { 'MMYY': lastSerial }
  const KEY_MASTER = 'axis_burgundy_ic_master_v2';
  const KEY_USER = 'axis_burgundy_ic_user_v2';
  const KEY_ROLE = 'axis_burgundy_ic_role_v2';

  function loadJSON(key, fallback) {
    try {
      const v = localStorage.getItem(key);
      return v ? JSON.parse(v) : fallback;
    } catch (e) {
      console.warn('Load failed', key, e);
      return fallback;
    }
  }
  function saveJSON(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (e) {
      toast(
        'error',
        'Storage Error',
        'Could not save locally. Storage may be full.',
      );
      return false;
    }
  }

  const HAS_FIREBASE = !!CONFIG.FIREBASE_URL;

  let state = {
    records: HAS_FIREBASE ? [] : loadJSON(KEY_RECORDS, []),

    audit: HAS_FIREBASE ? [] : loadJSON(KEY_AUDIT, []),

    draft: loadJSON(KEY_DRAFT, null),

    // Always start fresh session
    user: '',

    role: 'VIEWER',

    page: 'dashboard',

    editingRef: null,
  };

  /* =========================================================
               4. REAL-TIME SYNC LAYER
               Strategy:
                 - Firebase REST API + Server-Sent Events for live Pan-India sync
                 - BroadcastChannel for instant same-device cross-tab updates
                 - Storage events as additional cross-tab safety net
                 - Polling fallback if SSE blocked by network/proxy
            ========================================================= */
  const Sync = (function () {
    let bc = null;
    let es = null;
    let pollTimer = null;
    let online = false;

    function setStatus(label, kind) {
      const el = document.getElementById('syncStatus');
      if (!el) return;
      el.textContent = label;
      el.className = 'badge ' + (kind || 'offline');
      el.style.marginTop = '2px';
    }

    function broadcast(type, payload) {
      try {
        if (bc)
          bc.postMessage({ type, payload, src: state.user, ts: Date.now() });
      } catch (e) {}
    }

    function setupChannel() {
      if (typeof BroadcastChannel === 'undefined') return;
      bc = new BroadcastChannel(CONFIG.CHANNEL);
      bc.onmessage = (e) => {
        if (!e.data || e.data.src === state.user) return;
        handleRemoteEvent(e.data.type, e.data.payload);
      };
    }

    function setupStorageEvent() {
      window.addEventListener('storage', (e) => {
        if (e.key === KEY_RECORDS && e.newValue) {
          try {
            state.records = JSON.parse(e.newValue);
            refreshAll();
          } catch (err) {}
        }
        if (e.key === KEY_MASTER && e.newValue) {
          try {
            applyMaster(JSON.parse(e.newValue), false);
          } catch (err) {}
        }
      });
    }

    function handleRemoteEvent(type, payload) {
      if (type === 'record_upsert' && payload) {
        const idx = state.records.findIndex((r) => r.ref === payload.ref);
        if (idx >= 0) state.records[idx] = payload;
        else state.records.unshift(payload);
        saveJSON(KEY_RECORDS, state.records);
        refreshAll();
        if (state.role === 'ADMIN') {
          toast(
            'info',
            'Live update',
            'New/updated entry from ' +
              (payload.submittedBy || 'IC') +
              ' · ' +
              payload.ref,
          );
        }
      } else if (type === 'record_delete' && payload) {
        state.records = state.records.filter((r) => r.ref !== payload.ref);
        saveJSON(KEY_RECORDS, state.records);
        refreshAll();
      } else if (type === 'master_update' && payload) {
        applyMaster(payload, false);
        toast(
          'info',
          'Master refreshed',
          'Latest Master File data loaded from another session.',
        );
      }
    }

    /* ---------- Firebase REST helpers ---------- */
    function fbURL(path) {
      const base = (CONFIG.FIREBASE_URL || '').replace(/\/+$/, '');
      return base ? base + '/' + path.replace(/^\/+/, '') : '';
    }

    async function fbPut(path, value) {
      if (!CONFIG.FIREBASE_URL) return null;
      try {
        const r = await fetch(fbURL(path) + '?print=silent', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(value),
        });
        return r.ok;
      } catch (e) {
        console.warn('Firebase PUT failed', e);
        setStatus('Offline', 'offline');
        return false;
      }
    }
    async function fbDelete(path) {
      if (!CONFIG.FIREBASE_URL) return null;
      try {
        const r = await fetch(fbURL(path), { method: 'DELETE' });
        return r.ok;
      } catch (e) {
        return false;
      }
    }
    async function fbGet(path) {
      if (!CONFIG.FIREBASE_URL) return null;
      try {
        const r = await fetch(fbURL(path) + '.json');
        if (!r.ok) return null;
        return await r.json();
      } catch (e) {
        return null;
      }
    }

    /* Push a record to Firebase + broadcast locally */
    function pushRecord(record) {
      broadcast('record_upsert', record);
      if (CONFIG.FIREBASE_URL) {
        fbPut('ic/records/' + safeKey(record.ref) + '.json', record).then(
          (ok) => {
            if (ok) setStatus('Live', 'live');
          },
        );
      }
    }
    function deleteRecord(ref) {
      broadcast('record_delete', { ref });
      if (CONFIG.FIREBASE_URL) fbDelete('ic/records/' + safeKey(ref) + '.json');
    }
    function pushMaster(masterObj) {
      broadcast('master_update', masterObj);
      if (CONFIG.FIREBASE_URL) fbPut('ic/master.json', masterObj);
    }

    /* ---------- Server-Sent Events subscription for live remote updates ---------- */
    function subscribeSSE() {
      if (!CONFIG.FIREBASE_URL) return;
      try {
        es = new EventSource(fbURL('ic/records') + '.json');
        es.addEventListener('put', (ev) => handleFirebaseEvent(ev));
        es.addEventListener('patch', (ev) => handleFirebaseEvent(ev));
        es.onopen = () => {
          online = true;
          setStatus('Live', 'live');
        };
        es.onerror = () => {
          online = false;
          setStatus('Reconnect…', 'offline');
        };

        // Separate stream for master file updates
        const esMaster = new EventSource(fbURL('ic/master') + '.json');
        esMaster.addEventListener('put', (ev) => {
          try {
            const data = JSON.parse(ev.data);
            if (data && data.data) applyMaster(data.data, false);
          } catch (e) {}
        });
      } catch (e) {
        console.warn('SSE setup failed', e);
        startPolling();
      }
    }

    function handleFirebaseEvent(ev) {
      try {
        const evd = JSON.parse(ev.data);
        if (!evd) return;
        // PUT on /records.json → full replace; PUT on /records/<id>.json → single record
        if (evd.path === '/' && evd.data && typeof evd.data === 'object') {
          const remoteRecords = Object.values(evd.data).filter(Boolean);
          mergeRemoteRecords(remoteRecords);
        } else if (evd.path && evd.path !== '/') {
          const ref = evd.path.replace(/^\//, '');
          if (evd.data === null) {
            state.records = state.records.filter((r) => safeKey(r.ref) !== ref);
          } else if (evd.data && evd.data.ref) {
            const idx = state.records.findIndex((r) => r.ref === evd.data.ref);
            if (idx >= 0) state.records[idx] = evd.data;
            else state.records.unshift(evd.data);
          }
          saveJSON(KEY_RECORDS, state.records);
          refreshAll();
        }
      } catch (e) {
        console.warn('SSE parse error', e);
      }
    }

    function mergeRemoteRecords(remoteRecords) {
      if (!Array.isArray(remoteRecords)) return;
      const map = {};
      remoteRecords.forEach((r) => {
        if (r && r.ref) map[r.ref] = r;
      });
      state.records.forEach((r) => {
        if (!map[r.ref]) map[r.ref] = r;
      });
      state.records = Object.values(map).sort(
        (a, b) => new Date(b.submittedAt || 0) - new Date(a.submittedAt || 0),
      );
      saveJSON(KEY_RECORDS, state.records);
      refreshAll();
    }

    function startPolling() {
      if (pollTimer) clearInterval(pollTimer);
      pollTimer = setInterval(async () => {
        const remote = await fbGet('ic/records');
        if (remote && typeof remote === 'object') {
          mergeRemoteRecords(Object.values(remote));
          setStatus('Polling', 'live');
        }
      }, CONFIG.POLL_MS);
    }

    /* Initial pull from Firebase on app start */
    async function initialPull() {
      if (!CONFIG.FIREBASE_URL) return;
      const remoteRecords = await fbGet('ic/records');
      if (remoteRecords && typeof remoteRecords === 'object') {
        mergeRemoteRecords(Object.values(remoteRecords));
      }
      const remoteMaster = await fbGet('ic/master');
      if (
        remoteMaster &&
        typeof remoteMaster === 'object' &&
        remoteMaster.icMembers
      ) {
        applyMaster(remoteMaster, false);
      }
    }

    let _initialised = false;
    function init() {
      // Tear down any previous SSE connections so we can reconfigure on the fly
      try {
        if (es) {
          es.close();
          es = null;
        }
      } catch (e) {}
      try {
        if (pollTimer) {
          clearInterval(pollTimer);
          pollTimer = null;
        }
      } catch (e) {}
      if (!_initialised) {
        setupChannel();
        setupStorageEvent();
        _initialised = true;
      }
      if (CONFIG.FIREBASE_URL) {
        setStatus('Connecting…', 'offline');
        initialPull().then(() => {
          subscribeSSE();
        });
      } else {
        setStatus('Local', 'offline');
      }
    }

    return { init, pushRecord, deleteRecord, pushMaster, setStatus };
  })();

  /* Strip characters Firebase forbids in keys (. # $ [ ] /) */
  function safeKey(s) {
    return String(s || '').replace(/[.#$\[\]\/]/g, '_');
  }

  /* =========================================================
               5. UTILITIES
            ========================================================= */
  function $(id) {
    return document.getElementById(id);
  }
  function pad2(n) {
    return String(n).padStart(2, '0');
  }

  function fmtDate(dateStr) {
    if (!dateStr) return '';

    // Handle YYYY-MM-DD safely without timezone issues
    const parts = dateStr.split('-');

    if (parts.length === 3) {
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }

    const d = new Date(dateStr);

    if (isNaN(d)) return '';

    return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
  }

  function fmtDateTime(dateStr) {
    if (!dateStr) return '';

    const d = new Date(dateStr);

    if (isNaN(d)) return '';

    return (
      `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()} ` +
      `${pad2(d.getHours())}:${pad2(d.getMinutes())}`
    );
  }
  function escapeHtml(s) {
    if (s == null) return '';
    return String(s).replace(
      /[&<>"']/g,
      (ch) =>
        ({
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#39;',
        })[ch],
    );
  }
  function tsStamp() {
    const d = new Date();
    return (
      d.getFullYear() +
      String(d.getMonth() + 1).padStart(2, '0') +
      String(d.getDate()).padStart(2, '0') +
      '_' +
      String(d.getHours()).padStart(2, '0') +
      String(d.getMinutes()).padStart(2, '0')
    );
  }
  function safeName(s) {
    return (s || 'all').replace(/[^A-Za-z0-9]+/g, '_').slice(0, 40);
  }

  function syncRoleDropdownState() {
    const roleEl = $('currentRole');

    if (!roleEl) return;

    const hasUser = !!(state.user || '').trim();

    roleEl.disabled = !hasUser;

    // Optional visual reset
    if (!hasUser) {
      roleEl.value = 'VIEWER';
    }
  }

  /* =========================================================
               6. TOAST NOTIFICATIONS
            ========================================================= */
  function toast(kind, title, msg) {
    const c = $('toastContainer');
    if (!c) return;
    const t = document.createElement('div');
    t.className = 'toast ' + (kind || '');
    t.innerHTML =
      '<strong>' +
      escapeHtml(title) +
      '</strong><small>' +
      escapeHtml(msg || '') +
      '</small>';
    c.appendChild(t);
    setTimeout(() => {
      t.style.opacity = '0';
      t.style.transition = 'opacity .3s';
    }, 3500);
    setTimeout(() => t.remove(), 3900);
  }

  /* =========================================================
               7. AUDIT LOG
            ========================================================= */
  function logAudit(action, refId, detail) {
    const entry = {
      ts: new Date().toISOString(),
      user: state.user || '—',
      role: state.role,
      action: action,
      ref: refId || '—',
      detail: detail || '',
    };
    state.audit.unshift(entry);
    if (state.audit.length > 5000) state.audit.length = 5000;
    saveJSON(KEY_AUDIT, state.audit);
  }

  /* =========================================================
               8. NEW IC REFERENCE NUMBER  (MMYY_NNN, serial resets monthly)
               Per spec: short auto-generated format e.g. 0526_001
            ========================================================= */
  function nextRef() {
    const d = new Date();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yy = String(d.getFullYear()).slice(-2);
    const key = mm + yy;
    const seqMap = loadJSON(KEY_SEQ, {});
    // Determine highest existing serial for this month across all records.
    // Defensive: this handles two devices submitting in the same month even
    // before the latest sync has caught up, by scanning local state.records
    // (which includes anything we've pulled down from Firebase already).
    let max = seqMap[key] || 0;
    state.records.forEach((r) => {
      const m = (r.ref || '').match(/^(\d{4})_(\d+)/);
      if (m && m[1] === key) {
        const n = parseInt(m[2], 10);
        if (n > max) max = n;
      }
    });
    let serial = max + 1;
    let candidate = key + '_' + String(serial).padStart(3, '0');
    // If two ICs in different cities pick the same serial in the same second,
    // append a short device-disambiguating suffix on the trailing party only
    // when the candidate already exists locally (a rare race; the SSE/poll
    // merge will reconcile within seconds).
    while (state.records.some((r) => r.ref === candidate)) {
      serial++;
      candidate = key + '_' + String(serial).padStart(3, '0');
    }
    seqMap[key] = serial;
    saveJSON(KEY_SEQ, seqMap);
    return candidate;
  }

  /* =========================================================
               9. ROLE-BASED ACCESS CONTROL
            ========================================================= */
  function isAdminAuthorized(name) {
    if (!name) return false;
    const norm = name.trim().toLowerCase();
    // Direct match against the canonical allowlist
    if (CONFIG.ADMIN_ALLOWLIST.some((a) => a.trim().toLowerCase() === norm))
      return true;
    // Alias match — any alternate spelling registered for an admin
    const aliases = CONFIG.ADMIN_ALIASES || {};
    return Object.keys(aliases).some((canon) =>
      (aliases[canon] || []).some((alt) => alt.trim().toLowerCase() === norm),
    );
  }

  /* Return the canonical full name for any recognised admin spelling.
               Used so audit trails and headers always display one consistent name
               no matter which alias the user picked from a dropdown. */
  function canonicaliseAdmin(name) {
    if (!name) return name;
    const norm = name.trim().toLowerCase();
    if (CONFIG.ADMIN_ALLOWLIST.some((a) => a.trim().toLowerCase() === norm))
      return name.trim();
    const aliases = CONFIG.ADMIN_ALIASES || {};
    for (const canon of Object.keys(aliases)) {
      if (
        (aliases[canon] || []).some((alt) => alt.trim().toLowerCase() === norm)
      )
        return canon;
    }
    return name.trim();
  }

  function applyRolePermissions() {
    const role = state.role;
    // Tab visibility
    const newEntryTab = document.querySelector('.nav-tab[data-page="entry"]');
    if (newEntryTab)
      newEntryTab.style.display = role === 'VIEWER' ? 'none' : '';
    const masterTab = document.querySelector('.nav-tab[data-page="master"]');
    if (masterTab) masterTab.style.display = role === 'ADMIN' ? '' : 'none';
    // Submit disabled for viewers
    const btn = $('btnSubmit');
    if (btn) btn.disabled = role === 'VIEWER';
    // If viewer is currently on entry/master/records page, push them to dashboard
    if (
      role === 'VIEWER' &&
      (state.page === 'entry' || state.page === 'master')
    ) {
      go('dashboard');
    }
    if (role !== 'ADMIN' && state.page === 'master') {
      go('dashboard');
    }
    // IC Member auto-select in entry form (lock to current user)
    syncEntryFormToUser();
  }

  /* Records visible to current role:
                 ADMIN  → all records
                 IC_MEMBER → only records they submitted OR where member field = their username
                 VIEWER → no records (read-only dashboard view) */
  function getVisibleRecords() {
    if (state.role === 'ADMIN') return state.records;
    if (state.role === 'VIEWER') return [];
    // IC_MEMBER: own entries only
    const me = (state.user || '').trim().toLowerCase();
    if (!me) return [];
    return state.records.filter(
      (r) =>
        (r.submittedBy || '').trim().toLowerCase() === me ||
        (r.member || '').trim().toLowerCase() === me,
    );
  }

  /* Returns true if current role/user can edit a given record. */
  function canEdit(record) {
    if (state.role === 'VIEWER') return false;
    if (state.role === 'ADMIN') return true;
    if (state.role === 'IC_MEMBER') {
      const me = (state.user || '').trim().toLowerCase();
      return (
        (record.submittedBy || '').trim().toLowerCase() === me ||
        (record.member || '').trim().toLowerCase() === me
      );
    }
    return false;
  }

  function setUser(u) {
    const old = state.user;

    state.user = (u || '').trim();

    if (!state.user) {
      state.role = 'VIEWER';

      $('currentRole').value = 'VIEWER';
    }

    // Normalize admin aliases
    if (state.role === 'ADMIN' && state.user) {
      state.user = canonicaliseAdmin(state.user);
    }

    saveJSON(KEY_USER, state.user);

    // Validate admin access again
    if (state.role === 'ADMIN' && !isAdminAuthorized(state.user)) {
      alert('Unauthorized Admin Access');

      state.role = 'IC_MEMBER';

      saveJSON(KEY_ROLE, state.role);

      $('currentRole').value = 'IC_MEMBER';

      refreshUserSelect();

      toast(
        'error',
        'Unauthorized Admin Access',
        'Role reverted to IC Member.',
      );
    }

    // Sync dropdown UI
    if ($('currentUser')) {
      $('currentUser').value = state.user;
    }

    logAudit(
      'LOGIN',
      null,
      'User changed: ' + (old || '∅') + ' → ' + state.user,
    );

    applyRolePermissions();
    syncRoleDropdownState();
    refreshAll();
  }

  let pendingRole = null;
  const ADMIN_PASSWORD = 'qwer1234';

  function setRole(r) {
    const previous = state.role;

    // ADMIN protection
    if (r === 'ADMIN') {
      if (!state.user) {
        toast('error', 'Select User', 'Please select an admin user first.');

        $('currentRole').value = previous;

        return;
      }

      if (!isAdminAuthorized(state.user)) {
        alert('Unauthorized Admin Access');

        $('currentRole').value = previous;

        toast(
          'error',
          'Unauthorized Admin Access',
          'You are not on the authorised Admin roster.',
        );

        return;
      }

      pendingRole = r;

      openAdminPasswordModal(previous);

      return;
    }

    // ===== NON-ADMIN ROLE =====

    // ===== IC MEMBER PASSWORD CHECK =====

    if (r === 'IC_MEMBER') {
      if (!state.user) {
        toast('error', 'Select User', 'Please select a user first.');

        $('currentRole').value = previous;

        return;
      }

      openICPasswordModal(previous);

      return;
    }

    // ===== VIEWER ROLE =====

    state.role = r;

    saveJSON(KEY_ROLE, state.role);

    // Sync dropdown UI
    $('currentRole').value = r;

    // Refresh dropdowns
    refreshUserSelect();

    // Preserve selected user
    if ($('currentUser')) {
      $('currentUser').value = state.user || '';
    }

    // Apply RBAC permissions FIRST
    applyRolePermissions();

    logAudit('ROLE_CHANGE', null, 'Role set to ' + r);

    toast('info', 'Role updated', 'Active role: ' + r);

    // IMPORTANT:
    // refresh ONLY AFTER permissions fully updated
    refreshAll();
  }

  function openAdminPasswordModal(previousRole) {
    const modal = $('adminPasswordModal');

    modal.classList.add('show');

    $('adminPasswordInput').value = '';

    $('adminPasswordError').classList.remove('show');

    modal.dataset.previousRole = previousRole || 'IC_MEMBER';

    setTimeout(() => {
      $('adminPasswordInput').focus();
    }, 80);
  }

  function closeAdminPasswordModal(resetRole = true) {
    const modal = $('adminPasswordModal');

    modal.classList.remove('show');

    // Only reset dropdown if auth cancelled/failed
    if (resetRole) {
      const previous = modal.dataset.previousRole || 'IC_MEMBER';

      $('currentRole').value = previous;
    }

    pendingRole = null;

    $('adminPasswordInput').value = '';

    $('adminPasswordError').classList.remove('show');
  }

  function verifyAdminPassword() {
    const val = $('adminPasswordInput').value;

    if (val !== ADMIN_PASSWORD) {
      $('adminPasswordError').classList.add('show');

      return;
    }

    // SUCCESS
    state.role = 'ADMIN';

    // Normalize admin name
    state.user = canonicaliseAdmin(state.user);

    saveJSON(KEY_ROLE, state.role);

    saveJSON(KEY_USER, state.user);

    // IMPORTANT: update UI dropdowns FIRST
    $('currentRole').value = 'ADMIN';

    refreshUserSelect();

    $('currentUser').value = state.user;

    applyRolePermissions();

    logAudit('ROLE_CHANGE', null, 'Role set to ADMIN');

    toast('success', 'Access Granted', 'Admin access enabled');

    // false => don't revert dropdown
    closeAdminPasswordModal(false);

    refreshAll();
  }

  function openICPasswordModal(previousRole) {
    const modal = $('icPasswordModal');

    modal.classList.add('show');

    $('icPasswordInput').value = '';

    $('icPasswordError').classList.remove('show');

    modal.dataset.previousRole = previousRole || 'VIEWER';

    setTimeout(() => {
      $('icPasswordInput').focus();
    }, 80);
  }

  function closeICPasswordModal(resetRole = true) {
    const modal = $('icPasswordModal');

    modal.classList.remove('show');

    if (resetRole) {
      const previous = modal.dataset.previousRole || 'VIEWER';

      $('currentRole').value = previous;
    }

    $('icPasswordInput').value = '';

    $('icPasswordError').classList.remove('show');
  }

  function verifyICPassword() {
    const entered = $('icPasswordInput').value;

    const actualPassword = MASTER.icPasswords[state.user];

    if (!actualPassword || entered !== actualPassword) {
      $('icPasswordError').classList.add('show');

      return;
    }

    state.role = 'IC_MEMBER';

    saveJSON(KEY_ROLE, state.role);

    saveJSON(KEY_USER, state.user);

    $('currentRole').value = 'IC_MEMBER';

    refreshUserSelect();

    $('currentUser').value = state.user;

    applyRolePermissions();

    logAudit('ROLE_CHANGE', null, 'IC login successful');

    toast('success', 'Login Successful', 'Welcome ' + state.user);

    closeICPasswordModal(false);

    refreshAll();
  }

  function toggleAdminPassword() {
    const input = $('adminPasswordInput');

    input.type = input.type === 'password' ? 'text' : 'password';
  }

  function toggleICPassword() {
    const input = $('icPasswordInput');

    input.type = input.type === 'password' ? 'text' : 'password';
  }

  /* Populate the User dropdown based on selected role */
  function refreshUserSelect() {
    const sel = $('currentUser');

    if (!sel) return;

    // Same users for all roles
    let opts = MASTER.icMembers.slice();

    // Add admin allowlist users
    CONFIG.ADMIN_ALLOWLIST.forEach((a) => {
      if (!opts.includes(a)) {
        opts.push(a);
      }
    });

    // Remove duplicates
    opts = [...new Set(opts)];

    opts.sort((a, b) => a.localeCompare(b));

    sel.innerHTML =
      '<option value="">— Select user —</option>' +
      opts
        .map(
          (o) => `
      <option value="${escapeHtml(o)}">
        ${escapeHtml(o)}
      </option>
    `,
        )
        .join('');

    // Restore selected user
    sel.value = state.user || '';
  }

  /* When IC_MEMBER role is active, auto-fill the IC Member field in the entry form
               with the current user (if they match a master IC name) and lock it. */
  function syncEntryFormToUser() {
    const memberSel = $('f_member');
    if (!memberSel) return;
    if (
      state.role === 'IC_MEMBER' &&
      state.user &&
      MASTER.icMembers.includes(state.user)
    ) {
      memberSel.value = state.user;
      memberSel.disabled = true;
      onMemberChange();
    } else {
      memberSel.disabled = false;
    }
  }

  /* =========================================================
               10. MASTER DATA — LOAD / APPLY / UPLOAD / EXPORT
            ========================================================= */
  function loadMaster() {
    const stored = loadJSON(KEY_MASTER, null);
    MASTER =
      stored && stored.icMembers ? stored : structuredClone(DEFAULT_MASTER);
  }

  function applyMaster(masterObj, persist) {
    if (!masterObj || !masterObj.icMembers) return;
    // Always keep fixed lists (categories, meetingTypes, clientTypes, zones, admin allowlist)
    // overridable only if uploaded master explicitly provides them.
    MASTER = Object.assign({}, DEFAULT_MASTER, masterObj);
    if (persist) {
      saveJSON(KEY_MASTER, MASTER);
      Sync.pushMaster(MASTER);
    }
    populateMaster();
    renderMasterPage();
    refreshUserSelect();
    refreshAll();
  }

  function resetMaster() {
    if (
      !confirm('Restore the default Master data shipped with the application?')
    )
      return;
    localStorage.removeItem(KEY_MASTER);
    MASTER = structuredClone(DEFAULT_MASTER);
    MASTER = loadJSON(KEY_MASTER, DEFAULT_MASTER);
    loadMasterFromGoogleSheet();
    populateMaster();
    renderMasterPage();
    refreshUserSelect();
    refreshAll();
    Sync.pushMaster(MASTER);
    toast('info', 'Master restored', 'Default Master data has been reapplied.');
    logAudit('MASTER_RESET', null, 'Default master restored');
  }

  /* Parse uploaded Master Excel using SheetJS. Expected sheets:
                 RM_Data       : Region | RM_NAME | IC_NAME | RIC_NAME | …
                 IC_Zone_Map   : IC_Name | Zone
               Output: { icMembers, icToRic, icToZone, icToPartners, ricMembers }
            ========================================================= */
  function parseMasterWorkbook(workbook) {
    const ZONE_REMAP = {
      north: 'North-East Region',
      east: 'North-East Region',
      south: 'South Zone',
      west: 'West Zone',
      'north-east': 'North-East Region',
      'north-east region': 'North-East Region',
      'south zone': 'South Zone',
      'west zone': 'West Zone',
    };
    function bucketZone(region) {
      if (!region) return '';
      const k = String(region).trim().toLowerCase();
      // Try a full match first, then take first token (e.g., "North 1A" → "north")
      if (ZONE_REMAP[k]) return ZONE_REMAP[k];
      const head = k.split(/\s+/)[0];
      return ZONE_REMAP[head] || '';
    }

    /* Robust column lookup — accepts the original two-sheet layout
                   (RM_Data / IC_Zone_Map) as well as a single flat sheet with the
                   columns from the canonical Master_Excel_Tracker file:
                     SL | Region | RM_NAME | RM_EMP_ID | IC_NAME | IC_EMP_ID | RIC_NAME | confirmation
                   Header lookup is case- and whitespace-insensitive. */
    function pick(row, ...names) {
      if (!row) return '';
      const keys = Object.keys(row);
      for (const want of names) {
        const w = want.toLowerCase().replace(/[^a-z0-9]/g, '');
        for (const k of keys) {
          if (k.toLowerCase().replace(/[^a-z0-9]/g, '') === w) {
            const v = row[k];
            return v === null || v === undefined ? '' : String(v).trim();
          }
        }
      }
      return '';
    }
    function isJunk(s) {
      if (!s) return true;
      const t = String(s).trim().toLowerCase();
      return (
        !t ||
        t === '0' ||
        t === '-' ||
        t === 'nan' ||
        t === 'null' ||
        t === 'na' ||
        t === 'n/a'
      );
    }

    const out = {
      icToRic: {},
      icToZone: {},
      icToPartners: {},
      ricMembers: new Set(),
    };

    /* ---- Primary RM data sheet: prefer 'RM_Data', else first sheet ---- */
    const rmSheet =
      workbook.Sheets['RM_Data'] ||
      workbook.Sheets['Sheet1'] ||
      workbook.Sheets[workbook.SheetNames[0]];
    if (rmSheet) {
      const rows = XLSX.utils.sheet_to_json(rmSheet, { defval: '' });
      rows.forEach((row) => {
        const ic = pick(row, 'IC_NAME', 'IC_Name', 'IC Name', 'IC');
        const rm = pick(row, 'RM_NAME', 'RM_Name', 'RM Name', 'RM', 'Partner');
        const ric = pick(row, 'RIC_NAME', 'RIC_Name', 'RIC Name', 'RIC');
        const region = pick(row, 'Region', 'Zone');
        if (isJunk(ic)) return;
        // RIC mapping
        if (!isJunk(ric)) {
          out.icToRic[ic] = ric;
          out.ricMembers.add(ric);
        }
        // Zone (bucketed)
        const z = bucketZone(region);
        if (z && !out.icToZone[ic]) out.icToZone[ic] = z;
        // Partner mapping
        if (!isJunk(rm)) {
          if (!out.icToPartners[ic]) out.icToPartners[ic] = [];
          if (!out.icToPartners[ic].includes(rm)) out.icToPartners[ic].push(rm);
        } else if (!out.icToPartners[ic]) {
          out.icToPartners[ic] = [];
        }
      });
    }

    /* ---- IC_Zone_Map sheet (preferred zone source if present) ---- */
    const zoneSheet = workbook.Sheets['IC_Zone_Map'];
    if (zoneSheet) {
      const rows = XLSX.utils.sheet_to_json(zoneSheet, { defval: '' });
      rows.forEach((row) => {
        const ic = pick(row, 'IC_Name', 'IC_NAME', 'IC Name');
        const z = bucketZone(pick(row, 'Zone', 'Region'));
        if (ic && z) out.icToZone[ic] = z;
      });
    }

    // Sort partners alphabetically per IC
    Object.keys(out.icToPartners).forEach((ic) => out.icToPartners[ic].sort());

    const icMembers = Array.from(
      new Set(
        Object.keys(out.icToPartners)
          .concat(Object.keys(out.icToZone))
          .concat(Object.keys(out.icToRic)),
      ),
    ).sort();

    return {
      icMembers: icMembers,
      icToRic: out.icToRic,
      icToZone: out.icToZone,
      icToPartners: out.icToPartners,
      ricMembers: Array.from(out.ricMembers).sort(),
    };
  }

  function handleMasterUpload(ev) {
    if (state.role !== 'ADMIN') {
      toast(
        'error',
        'Permission denied',
        'Only Admins can upload Master files.',
      );
      return;
    }
    const file = ev.target.files && ev.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const parsed = parseMasterWorkbook(wb);
        if (!parsed.icMembers.length) {
          toast(
            'error',
            'Empty master',
            'No IC members found. Check sheet names and column headers.',
          );
          return;
        }
        applyMaster(parsed, true);
        $('masterStatus').textContent =
          'Last upload: ' +
          new Date().toLocaleString('en-IN') +
          ' · ' +
          parsed.icMembers.length +
          ' ICs';
        toast(
          'info',
          'Master updated',
          parsed.icMembers.length +
            ' IC members loaded · ' +
            Object.values(parsed.icToPartners).reduce(
              (s, a) => s + a.length,
              0,
            ) +
            ' partner mappings.',
        );
        logAudit('MASTER_UPLOAD', null, 'Master file uploaded: ' + file.name);
        ev.target.value = '';
      } catch (err) {
        console.error(err);
        toast(
          'error',
          'Upload failed',
          'Could not parse Master Excel: ' + err.message,
        );
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /* Export the current MASTER object as an Excel file (mirrors original schema) */
  function exportCurrentMaster() {
    const wb = XLSX.utils.book_new();
    // RM_Data sheet — one row per IC→Partner pair
    const rmRows = [
      [
        'Region',
        'RM_NAME',
        'IC_NAME',
        'IC_EMP_ID',
        'RIC_NAME',
        'confirmation',
        'Resigned',
        'Notes',
      ],
    ];
    MASTER.icMembers.forEach((ic) => {
      const zone = MASTER.icToZone[ic] || '';
      const ric = MASTER.icToRic[ic] || '';
      const partners = MASTER.icToPartners[ic] || [];
      if (!partners.length) {
        rmRows.push([zone, '', ic, '', ric, '', '', '']);
      } else {
        partners.forEach((p) =>
          rmRows.push([zone, p, ic, '', ric, '', '', '']),
        );
      }
    });
    const ws1 = XLSX.utils.aoa_to_sheet(rmRows);
    ws1['!cols'] = [
      { wch: 20 },
      { wch: 30 },
      { wch: 24 },
      { wch: 12 },
      { wch: 26 },
      { wch: 30 },
      { wch: 12 },
      { wch: 30 },
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'RM_Data');

    // IC_Zone_Map sheet
    const zRows = [['IC_Name', 'Zone']];
    MASTER.icMembers.forEach((ic) =>
      zRows.push([ic, MASTER.icToZone[ic] || '']),
    );
    const ws2 = XLSX.utils.aoa_to_sheet(zRows);
    ws2['!cols'] = [{ wch: 26 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'IC_Zone_Map');

    XLSX.writeFile(wb, 'Burgundy_Master_' + tsStamp() + '.xlsx');
    toast(
      'info',
      'Master exported',
      'Current master data downloaded as Excel.',
    );
  }

  /* =========================================================
               11. DROPDOWN POPULATION (driven by live MASTER)
            ========================================================= */
  function fillSelect(el, list, includeBlank, blankLabel) {
    if (!el) return;
    el.innerHTML =
      (includeBlank
        ? '<option value="">' +
          escapeHtml(blankLabel || '— Select —') +
          '</option>'
        : '') +
      list
        .map(
          (v) =>
            '<option value="' +
            escapeHtml(v) +
            '">' +
            escapeHtml(v) +
            '</option>',
        )
        .join('');
  }

  function populateMaster() {
    // Entry form selects
    fillSelect($('f_zone'), MASTER.zones, true);
    fillSelect($('f_partner'), [], true, 'Select IC Member first');
    fillSelect($('f_member'), MASTER.icMembers, true);
    fillSelect($('f_category'), MASTER.categories, true);
    fillSelect($('f_meetingType'), MASTER.meetingTypes, true);
    fillSelect($('f_clientType'), MASTER.clientTypes, true);

    // Filter selects
    fillSelect($('filt_zone'), MASTER.zones, true, 'All Zones');
    // Filter "Partner" uses union of all partners across ICs (sorted)
    const allPartners = Array.from(
      new Set(
        [].concat(
          ...MASTER.icMembers.map((ic) => MASTER.icToPartners[ic] || []),
        ),
      ),
    ).sort();
    fillSelect($('filt_partner'), allPartners, true, 'All Partners');
    fillSelect($('filt_member'), MASTER.icMembers, true, 'All Members');
    fillSelect($('filt_meetingType'), MASTER.meetingTypes, true, 'All');
    fillSelect($('filt_ric'), MASTER.ricMembers, true, 'All');

    // Report selects
    fillSelect($('rep_zone'), MASTER.zones, true, '— Select Zone —');
    fillSelect($('rep_partner'), allPartners, true, '— Select Partner —');
    refreshReportUserSelect();
  }

  function refreshReportUserSelect() {
    let users;
    if (state.role === 'ADMIN') {
      // Admin sees every submitter who has ever logged a record
      users = Array.from(
        new Set([state.user, ...state.records.map((r) => r.submittedBy)]),
      ).filter(Boolean);
    } else if (state.role === 'IC_MEMBER') {
      // IC member only sees themselves — never other ICs' names
      users = state.user ? [state.user] : [];
    } else {
      // Viewer has no per-user export
      users = [];
    }
    fillSelect($('rep_user'), users, true, '— Select User —');
  }

  /* When an IC Member is picked: auto-fill RIC, filter Zone to that IC's zone,
               filter Partner dropdown to only that IC's mapped partners. */
  function onMemberChange() {
    const m = $('f_member').value;
    // RIC
    $('f_ric').value = m && MASTER.icToRic[m] ? MASTER.icToRic[m] : '';
    // Zone — single zone per IC under the new 3-zone scheme
    const zoneEl = $('f_zone');
    if (m && MASTER.icToZone[m]) {
      zoneEl.innerHTML =
        '<option value="">— Select Zone —</option>' +
        '<option value="' +
        escapeHtml(MASTER.icToZone[m]) +
        '" selected>' +
        escapeHtml(MASTER.icToZone[m]) +
        '</option>';
    } else {
      fillSelect(zoneEl, MASTER.zones, true);
    }
    // Partner — strictly limited to mapped partners for this IC
    const partnerEl = $('f_partner');
    const hint = $('partnerHint');
    if (m && MASTER.icToPartners[m] && MASTER.icToPartners[m].length) {
      fillSelect(partnerEl, MASTER.icToPartners[m], true, 'Select Partner');
      if (hint)
        hint.textContent =
          MASTER.icToPartners[m].length + ' partner(s) mapped to this IC';
    } else {
      fillSelect(partnerEl, [], true, 'No partners mapped');
      if (hint)
        hint.textContent = 'No partners mapped to this IC in the Master file';
    }
  }

  /* =========================================================
               12. FORM HELPERS
            ========================================================= */
  const REQUIRED_FIELDS = [
    'f_date',
    'f_member',
    'f_zone',
    'f_partner',
    'f_meetingType',
    'f_category',
    'f_client',
    'f_clientType',
  ];

  function validate() {
    let ok = true;
    REQUIRED_FIELDS.forEach((id) => {
      const el = $(id);
      const group = el.closest('.form-group');
      const v = (el.value || '').toString().trim();
      if (!v) {
        group.classList.add('error');
        ok = false;
      } else {
        group.classList.remove('error');
      }
    });
    return ok;
  }

  let draftTimer = null;
  function startDraftAutoSave() {
    if (draftTimer) clearInterval(draftTimer);
    draftTimer = setInterval(() => {
      if (state.page === 'entry' && isFormDirty() && !state.editingRef)
        saveDraft(false);
    }, 15000);
  }
  function isFormDirty() {
    return (
      REQUIRED_FIELDS.some(
        (id) => ($(id).value || '').toString().trim() !== '',
      ) || ($('f_remarks').value || '').trim() !== ''
    );
  }
  function readForm() {
    const member = $('f_member').value;
    return {
      ref: $('f_ref').value,
      date: $('f_date').value,
      member: member,
      ric: member && MASTER.icToRic[member] ? MASTER.icToRic[member] : '',
      zone: $('f_zone').value,
      partner: $('f_partner').value,
      meetingType: $('f_meetingType').value,
      category: $('f_category').value,
      client: $('f_client').value.trim(),
      clientType: $('f_clientType').value,
      meetingMode: $('f_meetingMode').value,
      remarks: $('f_remarks').value.trim(),
    };
  }
  function writeForm(d) {
    if (!d) d = {};
    $('f_ref').value = d.ref || '';
    $('f_date').value = d.date || '';
    $('f_member').value = d.member || '';
    // Trigger member change to repopulate dependent dropdowns
    if (d.member) onMemberChange();
    $('f_ric').value = d.ric || (d.member && MASTER.icToRic[d.member]) || '';
    $('f_zone').value = d.zone || '';
    $('f_partner').value = d.partner || '';
    $('f_meetingType').value = d.meetingType || '';
    $('f_category').value = d.category || '';
    $('f_client').value = d.client || '';
    $('f_clientType').value = d.clientType || '';
    $('f_meetingMode').value = d.meetingMode || '';
    $('f_remarks').value = d.remarks || '';
  }
  function saveDraft(announce) {
    if (state.editingRef) return; // never auto-save over an edit
    const d = readForm();
    d.savedAt = new Date().toISOString();
    state.draft = d;
    saveJSON(KEY_DRAFT, d);
    $('draftStatus').textContent =
      'Draft saved · ' + new Date().toLocaleTimeString('en-IN');
    if (announce)
      toast('info', 'Draft saved', 'Your entry has been preserved locally.');
  }
  function loadDraft() {
    if (state.draft && !state.editingRef) {
      writeForm(state.draft);
      $('draftStatus').textContent =
        'Draft restored · ' +
        new Date(state.draft.savedAt).toLocaleTimeString('en-IN');
    }
  }
  function clearDraft() {
    state.draft = null;
    localStorage.removeItem(KEY_DRAFT);
    $('draftStatus').textContent = 'Not saved';
  }
  function clearForm() {
    document
      .querySelectorAll('#icForm .form-group.error')
      .forEach((g) => g.classList.remove('error'));
    writeForm({});
    clearDraft();
    if (state.editingRef) cancelEdit();
    toast('info', 'Form cleared', 'Discarded any unsaved changes.');
  }

  /* =========================================================
               13. SUBMIT / EDIT / UPDATE ENTRY
            ========================================================= */
  function submitEntry(e) {
    e.preventDefault();
    if (state.role === 'VIEWER')
      return toast(
        'error',
        'Permission denied',
        'Viewers cannot submit IC entries.',
      );
    if (!state.user)
      return toast(
        'error',
        'User required',
        'Please select your user name in the header before submitting.',
      );
    if (!validate())
      return toast(
        'error',
        'Validation failed',
        'Please complete all mandatory fields.',
      );

    // Editing existing entry?
    if (state.editingRef) {
      const idx = state.records.findIndex((r) => r.ref === state.editingRef);
      if (idx < 0)
        return toast(
          'error',
          'Not found',
          'Original record could not be located.',
        );
      const orig = state.records[idx];
      if (!canEdit(orig))
        return toast(
          'error',
          'Permission denied',
          'You may only edit your own entries.',
        );
      const updated = Object.assign({}, orig, readForm(), {
        ref: orig.ref, // preserve original ref
        submittedBy: orig.submittedBy, // preserve original creator
        submittedAt: orig.submittedAt, // preserve original timestamp
        editedBy: state.user,
        editedAt: new Date().toISOString(),
      });
      state.records[idx] = updated;
      saveJSON(KEY_RECORDS, state.records);
      Sync.pushRecord(updated);
      logAudit('UPDATE', updated.ref, 'Entry updated · ' + updated.client);
      cancelEdit();
      toast('info', 'Entry updated', 'Reference: ' + updated.ref);
      refreshAll();
      go('records');
      return;
    }

    // New entry
    const d = readForm();
    // Duplicate guard
    const dupe = state.records.find(
      (r) =>
        r.client === d.client &&
        r.member === d.member &&
        r.date === d.date &&
        r.partner === d.partner &&
        Date.now() - new Date(r.submittedAt).getTime() < 60000,
    );
    if (dupe)
      return toast(
        'warn',
        'Possible duplicate',
        'A near-identical entry was submitted seconds ago (' + dupe.ref + ').',
      );

    d.ref = nextRef();
    d.submittedBy = state.user;
    d.submittedAt = new Date().toISOString();
    d.editedBy = null;
    d.editedAt = null;
    // Note: entries are editable; no locked flag is stored.
    state.records.unshift(d);
    saveJSON(KEY_RECORDS, state.records);
    Sync.pushRecord(d);
    logAudit(
      'CREATE',
      d.ref,
      'Entry submitted · ' + d.client + ' · ' + d.member,
    );
    clearDraft();
    writeForm({});
    toast('info', 'IC entry submitted', 'Reference: ' + d.ref);
    refreshReportUserSelect();
    refreshAll();
    go('records');
  }

  /* Start editing a record — populate the form, switch button text, jump to entry page */
  function editEntry(ref) {
    const r = state.records.find((x) => x.ref === ref);
    if (!r) return;
    if (!canEdit(r))
      return toast(
        'error',
        'Permission denied',
        'You may only edit your own entries.',
      );
    state.editingRef = ref;
    writeForm(r);
    $('btnSubmit').textContent = 'Update Entry';
    $('btnCancelEdit').style.display = '';
    $('formTitle').textContent = 'Editing IC Entry · ' + r.ref;
    $('entryPageTitle').textContent = 'Edit IC Entry';
    $('draftStatus').textContent = 'Editing existing entry';
    go('entry');
    toast(
      'info',
      'Editing entry',
      'Make your changes and click Update Entry to save.',
    );
  }
  function cancelEdit() {
    state.editingRef = null;
    $('btnSubmit').textContent = 'Submit Entry';
    $('btnCancelEdit').style.display = 'none';
    $('formTitle').textContent = 'IC Entry — Draft';
    $('entryPageTitle').textContent = 'New Investment Committee Entry';
    writeForm({});
  }

  /* =========================================================
               14. DASHBOARD
            ========================================================= */
  function renderDashboard() {
    const recs = getVisibleRecords();
    const kpi = $('kpiGrid');
    const total = recs.length;
    const now = new Date();
    const mtdStart =
      now.getFullYear() +
      '-' +
      String(now.getMonth() + 1).padStart(2, '0') +
      '-01';
    const today = now.toISOString().slice(0, 10);
    const mtd = recs.filter((r) => r.date >= mtdStart).length;
    const todayCount = recs.filter((r) => r.date === today).length;
    const uniqueClients = new Set(
      recs.map((r) => (r.client || '').trim().toLowerCase()).filter(Boolean),
    ).size;
    const uniqueMembers = new Set(recs.map((r) => r.member).filter(Boolean))
      .size;
    const activeZones = new Set(recs.map((r) => r.zone).filter(Boolean)).size;

    const rows = getVisibleRecords();

    // Unique partners covered in CURRENT MONTH only
    const currentMonth = new Date().toISOString().slice(0, 7);

    const monthlyPartners = new Set();

    rows.forEach((r) => {
      if (!r.partner || !r.date) return;

      const recordMonth = String(r.date).slice(0, 7);

      if (recordMonth === currentMonth) {
        monthlyPartners.add(r.partner.trim());
      }
    });

    const uniquePartnersPerMonth = monthlyPartners.size;
    const scopeLabel =
      state.role === 'ADMIN'
        ? 'All India'
        : state.role === 'IC_MEMBER'
          ? 'My entries'
          : 'Read-only';
    kpi.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Total Entries (${scopeLabel})</div><div class="kpi-value">${total}</div><div class="kpi-foot">All-time IC submissions</div></div>
    <div class="kpi-card green"><div class="kpi-label">This Month</div><div class="kpi-value">${mtd}</div><div class="kpi-foot">Entries logged month-to-date</div></div>
    <div class="kpi-card amber"><div class="kpi-label">Today</div><div class="kpi-value">${todayCount}</div><div class="kpi-foot">Entries logged today</div></div>
    <div class="kpi-card gold"><div class="kpi-label">Unique Clients</div><div class="kpi-value">${uniqueClients}</div><div class="kpi-foot">Distinct investors covered</div></div>
    <div class="kpi-card"><div class="kpi-label">IC Members Active</div><div class="kpi-value">${uniqueMembers}</div><div class="kpi-foot">Distinct ICs who have logged</div></div>
    <div class="kpi-card red">
      <div class="kpi-label">Unique Partners Covered Per Month</div>
      <div class="kpi-value">${uniquePartnersPerMonth}</div>
      <div class="kpi-foot">Distinct partners covered in current month</div>
    </div>
  `;

    // Recent table
    const tb = $('recentTable').querySelector('tbody');
    const recent = recs.slice(0, 7);
    tb.innerHTML = recent.length
      ? recent
          .map(
            (r) => `
        <tr onclick="App.showDetail('${r.ref}')" style="cursor:pointer;">
          <td class="mono">${r.ref}</td>
          <td>${fmtDate(r.date)}</td>
          <td>${escapeHtml(r.member)}</td>
          <td>${escapeHtml(r.ric || '')}</td>
          <td>${escapeHtml(r.zone)}</td>
          <td>${escapeHtml(r.partner)}</td>
          <td>${escapeHtml(r.meetingType || '')}</td>
          <td>${escapeHtml(r.clientType || '')}</td>
        </tr>`,
          )
          .join('')
      : '<tr><td colspan="8" class="empty-state">No IC entries yet</td></tr>';

    // Zone distribution
    const zoneCounts = {};
    recs.forEach((r) => (zoneCounts[r.zone] = (zoneCounts[r.zone] || 0) + 1));
    const topZones = Object.entries(zoneCounts).sort((a, b) => b[1] - a[1]);
    const maxZ = Math.max(1, ...topZones.map((t) => t[1]));
    $('zoneDist').innerHTML = topZones.length
      ? topZones
          .map(([z, c]) => {
            const pct = (c / maxZ) * 100;
            return `<div style="margin-bottom:8px;">
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px;">
        <span style="font-weight:600;">${escapeHtml(z)}</span><span style="color:var(--ink-500);">${c}</span>
      </div>
      <div style="background:var(--ink-100);height:6px;border-radius:3px;overflow:hidden;">
        <div style="background:var(--burgundy-700);height:100%;width:${pct}%;transition:width .4s;"></div>
      </div>
    </div>`;
          })
          .join('')
      : '<div style="color:var(--ink-500);font-size:12px;">No data yet</div>';

    // Partner distribution (top 5)
    const partnerCounts = {};
    recs.forEach(
      (r) => (partnerCounts[r.partner] = (partnerCounts[r.partner] || 0) + 1),
    );
    const top = Object.entries(partnerCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    const maxP = Math.max(1, ...top.map((t) => t[1]));
    $('partnerDist').innerHTML = top.length
      ? top
          .map(([p, c]) => {
            const pct = (c / maxP) * 100;
            return `<div style="margin-bottom:8px;">
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px;">
        <span style="font-weight:600;">${escapeHtml(p)}</span><span style="color:var(--ink-500);">${c}</span>
      </div>
      <div style="background:var(--ink-100);height:6px;border-radius:3px;overflow:hidden;">
        <div style="background:var(--gold-500);height:100%;width:${pct}%;"></div>
      </div>
    </div>`;
          })
          .join('')
      : '<div style="color:var(--ink-500);font-size:12px;">No data yet</div>';

    $('navRecCount').textContent = total;
  }

  /* =========================================================
               15. RECORDS TABLE + FILTERS
            ========================================================= */
  function getFilteredRecords() {
    const search = ($('filt_search').value || '').toLowerCase().trim();
    const from = $('filt_from').value;
    const to = $('filt_to').value;
    const zone = $('filt_zone').value;
    const partner = $('filt_partner').value;
    const meetingType = $('filt_meetingType').value;
    const member = $('filt_member').value;
    const ric = $('filt_ric').value;

    return getVisibleRecords().filter((r) => {
      if (
        search &&
        !(
          (r.ref || '').toLowerCase().includes(search) ||
          (r.client || '').toLowerCase().includes(search) ||
          (r.partner || '').toLowerCase().includes(search) ||
          (r.member || '').toLowerCase().includes(search)
        )
      )
        return false;
      if (from && r.date < from) return false;
      if (to && r.date > to) return false;
      if (zone && r.zone !== zone) return false;
      if (partner && r.partner !== partner) return false;
      if (meetingType && r.meetingType !== meetingType) return false;
      if (member && r.member !== member) return false;
      if (ric && r.ric !== ric) return false;
      return true;
    });
  }

  function renderRecords() {
    if (state.role === 'VIEWER') {
      $('filtCount').textContent = 'Read-only role';
      $('recordsTable').querySelector('tbody').innerHTML =
        '<tr><td colspan="12"><div class="empty-state"><div class="icon">∅</div>Viewer role has dashboard access only.</div></td></tr>';
      return;
    }
    const list = getFilteredRecords();
    $('filtCount').textContent =
      list.length + ' record' + (list.length === 1 ? '' : 's') + ' matched';
    const tb = $('recordsTable').querySelector('tbody');
    if (!list.length) {
      tb.innerHTML =
        '<tr><td colspan="12"><div class="empty-state"><div class="icon">∅</div>No records match the current filters.</div></td></tr>';
      return;
    }
    tb.innerHTML = list
      .map(
        (r) => `
    <tr class="${r.editedAt ? 'edited' : ''}">
      <td class="mono">${r.ref}</td>
      <td>${fmtDate(r.date)}</td>
      <td>${escapeHtml(r.member)}</td>
      <td>${escapeHtml(r.ric || '')}</td>
      <td>${escapeHtml(r.zone)}</td>
      <td>${escapeHtml(r.partner)}</td>
      <td>${escapeHtml(r.category)}</td>
      <td>${escapeHtml(r.meetingType || '')}</td>
      <td>${escapeHtml(r.client)}</td>
      <td>${escapeHtml(r.clientType || '')}</td>
      <td>${escapeHtml(r.meetingMode || '')}</td>
      <td>${r.editedAt ? '<span class="badge edited">Edited</span>' : '<span class="badge approved">Submitted</span>'}</td>
      <td>
        <div class="row-actions">
          <button class="icon-btn" title="View" onclick="App.showDetail('${r.ref}')">👁</button>
          ${canEdit(r) ? `<button class="icon-btn" title="Edit" onclick="App.editEntry('${r.ref}')">✎</button>` : ''}
          <button class="icon-btn" title="Single export" onclick="App.exportOne('${r.ref}')">⤓</button>
          ${state.role === 'ADMIN' ? `<button class="icon-btn danger" title="Admin delete" onclick="App.adminDelete('${r.ref}')">✕</button>` : ''}
        </div>
      </td>
    </tr>`,
      )
      .join('');
  }

  function resetFilters() {
    [
      'filt_search',
      'filt_from',
      'filt_to',
      'filt_zone',
      'filt_partner',
      'filt_meetingType',
      'filt_member',
      'filt_ric',
    ].forEach((id) => ($(id).value = ''));
    renderRecords();
  }

  /* =========================================================
               16. DETAIL MODAL
            ========================================================= */
  function showDetail(ref) {
    const r = state.records.find((x) => x.ref === ref);
    if (!r) return;
    // IC members may only view their own records; viewers cannot open detail.
    if (state.role === 'VIEWER') {
      return toast(
        'error',
        'Permission denied',
        'Viewers cannot open record details.',
      );
    }
    if (state.role === 'IC_MEMBER' && !canEdit(r)) {
      return toast(
        'error',
        'Permission denied',
        'You may only view your own entries.',
      );
    }
    $('modalTitle').textContent = 'IC Record · ' + r.ref;
    const editedLine = r.editedAt
      ? `<div class="detail-item" style="grid-column:1/-1;"><div class="l">Last Edited</div><div class="v">${fmtDateTime(r.editedAt)} by ${escapeHtml(r.editedBy || '')}</div></div>`
      : '';
    $('modalBody').innerHTML = `
    <div class="detail-grid">
      <div class="detail-item"><div class="l">IC Reference</div><div class="v mono">${r.ref}</div></div>
      <div class="detail-item"><div class="l">IC Meeting Date</div><div class="v">${fmtDate(r.date)}</div></div>
      <div class="detail-item"><div class="l">IC Member</div><div class="v">${escapeHtml(r.member)}</div></div>
      <div class="detail-item"><div class="l">RIC Name</div><div class="v">${escapeHtml(r.ric || '')}</div></div>
      <div class="detail-item"><div class="l">Zone</div><div class="v">${escapeHtml(r.zone)}</div></div>
      <div class="detail-item"><div class="l">Partner Name</div><div class="v">${escapeHtml(r.partner)}</div></div>
      <div class="detail-item"><div class="l">Type of Meeting</div><div class="v">${escapeHtml(r.meetingType || '')}</div></div>
      <div class="detail-item"><div class="l">Product Category</div><div class="v">${escapeHtml(r.category)}</div></div>
      <div class="detail-item"><div class="l">Client</div><div class="v">${escapeHtml(r.client)}</div></div>
      <div class="detail-item"><div class="l">Type of Client</div><div class="v">${escapeHtml(r.clientType || '')}</div></div>
      <div class="detail-item"><div class="l">Submitted By</div><div class="v">${escapeHtml(r.submittedBy)}</div></div>
      <div class="detail-item"><div class="l">Submitted At</div><div class="v">${fmtDateTime(r.submittedAt)}</div></div>
      ${editedLine}
      <div class="detail-item" style="grid-column:1/-1;"><div class="l">Remarks / Outcomes</div><div class="v" style="white-space:pre-wrap;">${escapeHtml(r.remarks) || '<em style="color:var(--ink-500);">No remarks recorded</em>'}</div></div>
    </div>
    <div class="btn-row">
      ${canEdit(r) ? `<button class="btn btn-primary" onclick="App.closeModal();App.editEntry('${r.ref}')">Edit Entry</button>` : ''}
      <button class="btn btn-gold" onclick="App.exportOne('${r.ref}', 'xlsx')">Export this entry (Excel)</button>
      <button class="btn btn-ghost" onclick="App.exportOne('${r.ref}', 'csv')">CSV</button>
      <button class="btn btn-ghost" onclick="window.print()">Print</button>
    </div>`;
    $('modalBackdrop').classList.add('show');
    logAudit('VIEW', r.ref, 'Detail viewed');
  }
  function closeModal() {
    $('modalBackdrop').classList.remove('show');
  }

  /* =========================================================
               17. ADMIN DELETE
            ========================================================= */
  function adminDelete(ref) {
    if (state.role !== 'ADMIN')
      return toast(
        'error',
        'Permission denied',
        'Only Admins can delete records.',
      );
    if (
      !confirm(
        'Admin override: permanently delete entry ' +
          ref +
          '?\nThis action will be recorded in the audit trail.',
      )
    )
      return;
    const r = state.records.find((x) => x.ref === ref);
    state.records = state.records.filter((x) => x.ref !== ref);
    saveJSON(KEY_RECORDS, state.records);
    Sync.deleteRecord(ref);
    logAudit(
      'DELETE',
      ref,
      'Admin deletion · ' + (r ? r.client + ' · ' + r.member : ''),
    );
    refreshAll();
    toast('warn', 'Entry deleted', ref + ' has been removed (audit logged).');
  }

  /* =========================================================
               18. EXPORTS — column defs + builders
            ========================================================= */
  const EXPORT_COLS = [
    { k: 'ref', h: 'IC Reference No.' },
    { k: 'date', h: 'IC Meeting Date' },
    { k: 'member', h: 'IC Member' },
    { k: 'ric', h: 'RIC Name' },
    { k: 'zone', h: 'Zone' },
    { k: 'partner', h: 'Partner Name' },
    { k: 'meetingType', h: 'Type of Meeting' },
    { k: 'category', h: 'Product Category' },
    { k: 'client', h: 'Client / Investor' },
    { k: 'clientType', h: 'Type of Client' },
    { k: 'meetingMode', h: 'Mode of Meeting' },
    { k: 'remarks', h: 'Remarks / Outcomes' },
    { k: 'submittedBy', h: 'Submitted By' },
    { k: 'submittedAt', h: 'Submitted At (UTC)' },
    { k: 'editedBy', h: 'Last Edited By' },
    { k: 'editedAt', h: 'Last Edited At (UTC)' },
  ];

  function toRows(records) {
    return records.map((r) =>
      EXPORT_COLS.map((c) => {
        const v = r[c.k];
        return v == null ? '' : v;
      }),
    );
  }

  function buildSheet(records, title) {
    const headers = EXPORT_COLS.map((c) => c.h);
    const rows = toRows(records);
    const aoa = [
      ['AXIS BANK · BURGUNDY PRIVATE · INVESTMENT COMMITTEE'],
      [title || 'IC Records Report'],
      [
        'Generated: ' +
          new Date().toLocaleString('en-IN') +
          '  ·  By: ' +
          (state.user || '—') +
          '  ·  Records: ' +
          records.length,
      ],
      [],
      headers,
      ...rows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [
      { wch: 14 },
      { wch: 14 },
      { wch: 24 },
      { wch: 22 },
      { wch: 18 },
      { wch: 28 },
      { wch: 20 },
      { wch: 22 },
      { wch: 28 },
      { wch: 20 },
      { wch: 48 },
      { wch: 18 },
      { wch: 22 },
      { wch: 18 },
      { wch: 22 },
    ];
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: headers.length - 1 } },
    ];
    const headerRowIdx = 4;
    headers.forEach((_, c) => {
      const cell = XLSX.utils.encode_cell({ r: headerRowIdx, c });
      if (ws[cell]) {
        ws[cell].s = {
          fill: { fgColor: { rgb: '7A1432' } },
          font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
          alignment: { horizontal: 'left', vertical: 'center' },
        };
      }
    });
    [0, 1, 2].forEach((r) => {
      const cell = XLSX.utils.encode_cell({ r, c: 0 });
      if (ws[cell]) {
        ws[cell].s = {
          font: {
            bold: r < 2,
            sz: r === 0 ? 13 : r === 1 ? 12 : 10,
            color: { rgb: r === 2 ? '6B6363' : '4A0E22' },
          },
          alignment: { horizontal: 'left' },
        };
      }
    });
    return ws;
  }

  function downloadWorkbook(records, filename, sheetTitle, extras) {
    if (!records.length)
      return toast(
        'warn',
        'Nothing to export',
        'There are no records matching the criteria.',
      );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      buildSheet(records, sheetTitle),
      'IC Records',
    );
    if (extras && extras.length)
      extras.forEach((ex) => XLSX.utils.book_append_sheet(wb, ex.ws, ex.name));
    XLSX.writeFile(wb, filename);
    logAudit('EXPORT', '—', filename + ' · ' + records.length + ' records');
    toast('info', 'Export complete', filename);
  }

  function downloadCSV(records, filename, sheetTitle) {
    if (!records.length)
      return toast(
        'warn',
        'Nothing to export',
        'There are no records matching the criteria.',
      );
    const headers = EXPORT_COLS.map((c) => c.h);
    const rows = toRows(records);
    const csvEscape = (v) => {
      if (v == null) return '';
      const s = String(v);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const meta = [
      '"AXIS BANK · BURGUNDY PRIVATE · INVESTMENT COMMITTEE"',
      '"' + (sheetTitle || 'IC Records') + '"',
      '"Generated: ' +
        new Date().toLocaleString('en-IN') +
        ' by ' +
        (state.user || '—') +
        ' · ' +
        records.length +
        ' records"',
      '',
    ].join('\n');
    const body = [
      headers.map(csvEscape).join(','),
      ...rows.map((r) => r.map(csvEscape).join(',')),
    ].join('\n');
    const blob = new Blob(['\uFEFF' + meta + '\n' + body], {
      type: 'text/csv;charset=utf-8',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    logAudit(
      'EXPORT',
      '—',
      filename + ' · ' + records.length + ' records (CSV)',
    );
    toast('info', 'CSV export complete', filename);
  }

  /* =========================================================
               19. EXPORT FACADES (all respect role-based visibility)
            ========================================================= */
  function exportAll(fmt) {
    const recs = getVisibleRecords();
    const fn = 'IC_Master_Report_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') downloadCSV(recs, fn, 'Complete Master Report');
    else downloadWorkbook(recs, fn, 'Complete Master Report');
  }
  function exportFiltered(fmt) {
    const list = getFilteredRecords();
    const fn = 'IC_Filtered_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') downloadCSV(list, fn, 'Filtered IC Report');
    else downloadWorkbook(list, fn, 'Filtered IC Report');
  }
  function exportOne(ref, fmt) {
    fmt = fmt || 'xlsx';
    const r = state.records.find((x) => x.ref === ref);
    if (!r) return;
    if (state.role !== 'ADMIN' && !canEdit(r))
      return toast(
        'error',
        'Permission denied',
        'You cannot export entries that are not yours.',
      );
    const fn = 'IC_' + safeName(r.ref) + '.' + fmt;
    if (fmt === 'csv') downloadCSV([r], fn, 'Single IC Entry · ' + r.ref);
    else downloadWorkbook([r], fn, 'Single IC Entry · ' + r.ref);
  }
  function exportByDate(fmt) {
    const from = $('rep_from').value,
      to = $('rep_to').value;
    if (!from || !to)
      return toast(
        'warn',
        'Date range required',
        'Please pick both From and To dates.',
      );
    const list = getVisibleRecords().filter(
      (r) => r.date >= from && r.date <= to,
    );
    const fn = 'IC_DateRange_' + from + '_to_' + to + '.' + fmt;
    if (fmt === 'csv')
      downloadCSV(list, fn, 'Date Range · ' + from + ' to ' + to);
    else downloadWorkbook(list, fn, 'Date Range · ' + from + ' to ' + to);
  }
  function exportByZone(fmt) {
    const z = $('rep_zone').value;
    if (!z) return toast('warn', 'Zone required', 'Please choose a zone.');
    const list = getVisibleRecords().filter((r) => r.zone === z);
    const fn = 'IC_Zone_' + safeName(z) + '_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') downloadCSV(list, fn, 'Zone · ' + z);
    else downloadWorkbook(list, fn, 'Zone · ' + z);
  }
  function exportByPartner(fmt) {
    const p = $('rep_partner').value;
    if (!p)
      return toast('warn', 'Partner required', 'Please choose a partner.');
    const list = getVisibleRecords().filter((r) => r.partner === p);
    const fn = 'IC_Partner_' + safeName(p) + '_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') downloadCSV(list, fn, 'Partner · ' + p);
    else downloadWorkbook(list, fn, 'Partner · ' + p);
  }
  function exportByUser(fmt) {
    const u = $('rep_user').value;
    if (!u) return toast('warn', 'User required', 'Please choose a user.');
    const list = getVisibleRecords().filter((r) => r.submittedBy === u);
    const fn = 'IC_User_' + safeName(u) + '_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') downloadCSV(list, fn, 'User · ' + u);
    else downloadWorkbook(list, fn, 'User · ' + u);
  }

  /* MIS PACK — multi-sheet workbook with master + summaries + audit */
  function exportMISPack() {
    const recs = getVisibleRecords();
    if (!recs.length)
      return toast(
        'warn',
        'Nothing to export',
        'No records available for MIS pack.',
      );

    // Sheet 2: Summary by Zone
    const byZone = {};
    recs.forEach((r) => {
      if (!byZone[r.zone]) byZone[r.zone] = { count: 0 };
      MASTER.meetingTypes.forEach((mt) => {
        byZone[r.zone][mt] = byZone[r.zone][mt] || 0;
      });
      byZone[r.zone].count++;
      if (byZone[r.zone][r.meetingType] !== undefined)
        byZone[r.zone][r.meetingType]++;
    });
    const zoneRows = [
      ['Zone Summary'],
      [],
      ['Zone', 'Total Entries', ...MASTER.meetingTypes],
    ];
    Object.keys(byZone)
      .sort()
      .forEach((z) => {
        const s = byZone[z];
        zoneRows.push([
          z,
          s.count,
          ...MASTER.meetingTypes.map((mt) => s[mt] || 0),
        ]);
      });
    const wsZone = XLSX.utils.aoa_to_sheet(zoneRows);
    wsZone['!cols'] = [
      { wch: 24 },
      { wch: 14 },
      ...MASTER.meetingTypes.map(() => ({ wch: 18 })),
    ];

    // Sheet 3: Summary by Partner
    const byPartner = {};
    recs.forEach((r) => {
      if (!byPartner[r.partner])
        byPartner[r.partner] = { count: 0, clients: new Set() };
      byPartner[r.partner].count++;
      if (r.client)
        byPartner[r.partner].clients.add(r.client.trim().toLowerCase());
    });
    const partnerRows = [
      ['Partner Summary'],
      [],
      ['Partner', 'Entries', 'Unique Clients'],
    ];
    Object.entries(byPartner)
      .sort((a, b) => b[1].count - a[1].count)
      .forEach(([p, s]) => partnerRows.push([p, s.count, s.clients.size]));
    const wsPartner = XLSX.utils.aoa_to_sheet(partnerRows);
    wsPartner['!cols'] = [{ wch: 34 }, { wch: 12 }, { wch: 16 }];

    // Sheet 4: Pivot by RIC × Meeting Type
    const ricPivot = {};
    recs.forEach((r) => {
      const key = r.ric || '—';
      if (!ricPivot[key]) {
        ricPivot[key] = { Total: 0 };
        MASTER.meetingTypes.forEach((mt) => {
          ricPivot[key][mt] = 0;
        });
      }
      if (ricPivot[key][r.meetingType] !== undefined)
        ricPivot[key][r.meetingType]++;
      ricPivot[key].Total++;
    });
    const pivotRows = [
      ['RIC × Meeting-Type Pivot'],
      [],
      ['RIC', ...MASTER.meetingTypes, 'Total'],
    ];
    Object.entries(ricPivot).forEach(([ric, s]) =>
      pivotRows.push([
        ric,
        ...MASTER.meetingTypes.map((mt) => s[mt] || 0),
        s.Total,
      ]),
    );
    const wsPivot = XLSX.utils.aoa_to_sheet(pivotRows);
    wsPivot['!cols'] = [
      { wch: 26 },
      ...MASTER.meetingTypes.map(() => ({ wch: 18 })),
      { wch: 10 },
    ];

    // Sheet 5: Pivot by Product Category × Type of Client
    const catPivot = {};
    recs.forEach((r) => {
      catPivot[r.category] = catPivot[r.category] || {};
      MASTER.clientTypes.forEach((ct) => {
        catPivot[r.category][ct] = catPivot[r.category][ct] || 0;
      });
      catPivot[r.category].Total = catPivot[r.category].Total || 0;
      if (catPivot[r.category][r.clientType] != null)
        catPivot[r.category][r.clientType]++;
      catPivot[r.category].Total++;
    });
    const catRows = [
      ['Category × Client-Type Pivot'],
      [],
      ['Product Category', ...MASTER.clientTypes, 'Total'],
    ];
    Object.entries(catPivot).forEach(([cat, s]) =>
      catRows.push([
        cat,
        ...MASTER.clientTypes.map((ct) => s[ct] || 0),
        s.Total,
      ]),
    );
    const wsCat = XLSX.utils.aoa_to_sheet(catRows);
    wsCat['!cols'] = [
      { wch: 30 },
      ...MASTER.clientTypes.map(() => ({ wch: 18 })),
      { wch: 10 },
    ];

    // Sheet 6: Audit Log
    const auditRows = [
      ['Digital Audit Trail'],
      [],
      ['Timestamp', 'User', 'Role', 'Action', 'Ref', 'Detail'],
    ];
    state.audit.forEach((a) =>
      auditRows.push([a.ts, a.user, a.role, a.action, a.ref, a.detail]),
    );
    const wsAudit = XLSX.utils.aoa_to_sheet(auditRows);
    wsAudit['!cols'] = [
      { wch: 24 },
      { wch: 18 },
      { wch: 12 },
      { wch: 14 },
      { wch: 18 },
      { wch: 50 },
    ];

    downloadWorkbook(
      recs,
      'IC_MIS_Pack_' + tsStamp() + '.xlsx',
      'Complete Master Report',
      [
        { name: 'Zone Summary', ws: wsZone },
        { name: 'Partner Summary', ws: wsPartner },
        { name: 'RIC × Meeting Pivot', ws: wsPivot },
        { name: 'Category × Client', ws: wsCat },
        { name: 'Audit Trail', ws: wsAudit },
      ],
    );
  }

  /* Audit export */
  function exportAudit(fmt) {
    if (!state.audit.length)
      return toast('warn', 'Nothing to export', 'Audit log is empty.');
    const fn = 'IC_Audit_Trail_' + tsStamp() + '.' + fmt;
    if (fmt === 'csv') {
      const head = ['Timestamp', 'User', 'Role', 'Action', 'Ref', 'Detail'];
      const rows = state.audit.map((a) => [
        a.ts,
        a.user,
        a.role,
        a.action,
        a.ref,
        a.detail,
      ]);
      const csvEscape = (v) => {
        const s = String(v || '');
        return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      };
      const body = [
        head.join(','),
        ...rows.map((r) => r.map(csvEscape).join(',')),
      ].join('\n');
      const blob = new Blob(['\uFEFF' + body], {
        type: 'text/csv;charset=utf-8',
      });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = fn;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } else {
      const wb = XLSX.utils.book_new();
      const aoa = [
        ['AXIS BANK · BURGUNDY PRIVATE · IC AUDIT TRAIL'],
        [
          'Generated: ' +
            new Date().toLocaleString('en-IN') +
            ' · By: ' +
            (state.user || '—') +
            ' · Events: ' +
            state.audit.length,
        ],
        [],
        ['Timestamp', 'User', 'Role', 'Action', 'Ref', 'Detail'],
        ...state.audit.map((a) => [
          a.ts,
          a.user,
          a.role,
          a.action,
          a.ref,
          a.detail,
        ]),
      ];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [
        { wch: 24 },
        { wch: 18 },
        { wch: 12 },
        { wch: 14 },
        { wch: 18 },
        { wch: 50 },
      ];
      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
      ];
      XLSX.utils.book_append_sheet(wb, ws, 'Audit Trail');
      XLSX.writeFile(wb, fn);
    }
    logAudit('EXPORT', '—', 'Audit trail exported: ' + fn);
    toast('info', 'Audit exported', fn);
  }

  /* Email-ready text summary */
  function emailReadySummary() {
    const recs = getVisibleRecords();
    if (!recs.length)
      return toast('warn', 'Nothing to summarise', 'No records yet.');
    const total = recs.length;
    const uniqueClients = new Set(
      recs.map((r) => (r.client || '').trim().toLowerCase()).filter(Boolean),
    ).size;

    // Top 5 IC Members by entry count
    const memberCount = {};
    recs.forEach((r) => {
      memberCount[r.member] = (memberCount[r.member] || 0) + 1;
    });
    const top5 = Object.entries(memberCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    const lines = [
      'AXIS BANK · BURGUNDY PRIVATE',
      'Investment Committee — Status Update',
      'Generated: ' + new Date().toLocaleString('en-IN'),
      'Scope: ' +
        (state.role === 'ADMIN'
          ? 'All India'
          : (state.user || 'IC Member') + ' (personal)'),
      '',
      'KPI SUMMARY',
      '-----------------------',
      'Total Entries        : ' + total,
      'Unique Clients       : ' + uniqueClients,
      '',
      'BY MEETING TYPE',
      '-----------------------',
      ...MASTER.meetingTypes.map(
        (mt) =>
          mt +
          ' '.repeat(Math.max(1, 22 - mt.length)) +
          ': ' +
          recs.filter((r) => r.meetingType === mt).length,
      ),
      '',
      'BY CLIENT TYPE',
      '-----------------------',
      ...MASTER.clientTypes.map(
        (ct) =>
          ct +
          ' '.repeat(Math.max(1, 22 - ct.length)) +
          ': ' +
          recs.filter((r) => r.clientType === ct).length,
      ),
      '',
      'BY ZONE',
      '-----------------------',
      ...MASTER.zones.map(
        (z) =>
          z +
          ' '.repeat(Math.max(1, 22 - z.length)) +
          ': ' +
          recs.filter((r) => r.zone === z).length,
      ),
      '',
      'TOP 5 IC MEMBERS (by entry count)',
      '-----------------------',
      ...top5.map(
        (row, i) => i + 1 + '. ' + row[0] + ' · ' + row[1] + ' entries',
      ),
      '',
      '— End of summary —',
    ];
    const text = lines.join('\n');

    $('modalTitle').textContent = 'Email-Ready Summary';
    $('modalBody').innerHTML = `
    <textarea readonly style="width:100%;min-height:340px;font-family:var(--font-mono);font-size:12px;background:var(--paper);">${escapeHtml(text)}</textarea>
    <div class="btn-row">
      <button class="btn btn-primary" onclick="navigator.clipboard.writeText(this.parentElement.parentElement.querySelector('textarea').value);App.toast('info','Copied','Summary copied to clipboard.')">Copy to Clipboard</button>
      <button class="btn btn-ghost" onclick="App.closeModal()">Close</button>
    </div>`;
    $('modalBackdrop').classList.add('show');
    logAudit('EXPORT', '—', 'Email summary generated');
  }

  /* =========================================================
               20. JSON BACKUP (cross-session manual sync)
            ========================================================= */
  function exportDataJSON() {
    // Viewers may not export at all; IC members may only export their own records.
    if (state.role === 'VIEWER') {
      return toast('error', 'Permission denied', 'Viewers cannot export data.');
    }
    const exportRecords =
      state.role === 'ADMIN' ? state.records : getVisibleRecords();
    const auditScope =
      state.role === 'ADMIN'
        ? state.audit
        : state.audit.filter(
            (a) =>
              (a.user || '').trim().toLowerCase() ===
              (state.user || '').trim().toLowerCase(),
          );
    const payload = {
      _version: 3,
      _exported: new Date().toISOString(),
      _exportedBy: state.user,
      _scope: state.role,
      records: exportRecords,
      audit: auditScope,
      seq: loadJSON(KEY_SEQ, {}),
      master: state.role === 'ADMIN' ? MASTER : null, // master only embedded for admins
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: 'application/json',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'IC_Data_Backup_' + tsStamp() + '.json';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    logAudit(
      'EXPORT',
      '—',
      'Full data backup exported as JSON · scope=' + state.role,
    );
    toast(
      'info',
      'Data exported',
      'Share this JSON file with other users to sync records.',
    );
  }

  function importDataJSON() {
    // Only admins may merge an externally produced backup into the shared dataset.
    if (state.role !== 'ADMIN') {
      return toast(
        'error',
        'Permission denied',
        'Only Admins can import data backups.',
      );
    }
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json,application/json';
    input.onchange = function (e) {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function (ev) {
        try {
          const payload = JSON.parse(ev.target.result);
          if (!payload.records || !Array.isArray(payload.records)) {
            toast(
              'error',
              'Invalid file',
              'This file does not appear to be a valid IC data backup.',
            );
            return;
          }
          const importedRefs = new Set(payload.records.map((r) => r.ref));
          const localOnly = state.records.filter(
            (r) => !importedRefs.has(r.ref),
          );
          state.records = [...payload.records, ...localOnly].sort(
            (a, b) => new Date(b.submittedAt) - new Date(a.submittedAt),
          );
          saveJSON(KEY_RECORDS, state.records);
          if (payload.seq) saveJSON(KEY_SEQ, payload.seq);
          const auditTs = new Set(
            state.audit.map((a) => a.ts + a.action + a.ref),
          );
          const newAudit = (payload.audit || []).filter(
            (a) => !auditTs.has(a.ts + a.action + a.ref),
          );
          state.audit = [...state.audit, ...newAudit]
            .sort((a, b) => new Date(b.ts) - new Date(a.ts))
            .slice(0, 5000);
          saveJSON(KEY_AUDIT, state.audit);
          if (payload.master && payload.master.icMembers)
            applyMaster(payload.master, true);
          logAudit(
            'IMPORT',
            '—',
            'Data imported from ' +
              file.name +
              ' · ' +
              payload.records.length +
              ' records merged',
          );
          refreshAll();
          toast(
            'info',
            'Import successful',
            payload.records.length + ' records loaded from backup.',
          );
        } catch (err) {
          toast(
            'error',
            'Import failed',
            'Could not parse the JSON file: ' + err.message,
          );
        }
      };
      reader.readAsText(file);
    };
    document.body.appendChild(input);
    input.click();
    input.remove();
  }

  /* =========================================================
               21. MASTER FILE PAGE RENDER
            ========================================================= */
  function renderSyncConfigPanel() {
    const input = $('fbUrlInput');
    const status = $('syncConfigStatus');
    const hint = $('fbUrlHint');
    if (!input) return;
    const url = CONFIG.FIREBASE_URL || '';
    input.value = url;
    if (status) {
      status.textContent = url
        ? 'Configured · live sync'
        : 'Not configured · local-only';
      status.style.color = url ? 'var(--success)' : 'var(--ink-500)';
    }
    if (hint) {
      if (url) {
        hint.innerHTML =
          '✓ Connected to <span style="color:var(--burgundy-700);">' +
          escapeHtml(url) +
          '</span> — entries sync across all admins in real time.';
      } else {
        hint.textContent =
          'No Firebase URL set. Records are stored locally on this device only.';
      }
    }
  }

  function saveFirebaseURL() {
    if (state.role !== 'ADMIN') {
      toast(
        'error',
        'Permission denied',
        'Only Admins can change sync configuration.',
      );
      return;
    }
    const v = ($('fbUrlInput').value || '').trim();
    if (setFirebaseURL(v)) renderSyncConfigPanel();
  }

  function clearFirebaseURL() {
    if (state.role !== 'ADMIN') {
      toast(
        'error',
        'Permission denied',
        'Only Admins can change sync configuration.',
      );
      return;
    }
    if (
      !confirm(
        'Disconnect from real-time sync? Records will remain on this device but will no longer sync to other admins.',
      )
    )
      return;
    $('fbUrlInput').value = '';
    setFirebaseURL('');
    renderSyncConfigPanel();
  }

  function copyShareLink() {
    if (!CONFIG.FIREBASE_URL) {
      toast(
        'error',
        'Not configured',
        'Set a Firebase URL first, then generate a share link.',
      );
      return;
    }
    const base = window.location.href.split('?')[0].split('#')[0];
    const link = base + '?fb=' + encodeURIComponent(CONFIG.FIREBASE_URL);
    navigator.clipboard
      .writeText(link)
      .then(() => {
        toast(
          'info',
          'Share link copied',
          "Send this link to other admins — they'll auto-configure on first visit.",
        );
      })
      .catch(() => {
        prompt('Copy this share link:', link);
      });
  }

  function renderMasterPage() {
    renderSyncConfigPanel();
    const kpi = $('masterKpi');
    if (kpi) {
      const totalP = Object.values(MASTER.icToPartners).reduce(
        (s, a) => s + a.length,
        0,
      );
      kpi.innerHTML = `
      <div class="kpi-card"><div class="kpi-label">IC Members</div><div class="kpi-value">${MASTER.icMembers.length}</div><div class="kpi-foot">From Master File</div></div>
      <div class="kpi-card green"><div class="kpi-label">Partners Mapped</div><div class="kpi-value">${totalP}</div><div class="kpi-foot">Across all IC members</div></div>
      <div class="kpi-card gold"><div class="kpi-label">RICs</div><div class="kpi-value">${MASTER.ricMembers.length}</div><div class="kpi-foot">Regional IC Heads</div></div>
      <div class="kpi-card amber"><div class="kpi-label">Zones</div><div class="kpi-value">${MASTER.zones.length}</div><div class="kpi-foot">South / West / North-East</div></div>
    `;
    }
    const icTb =
      $('masterICTable') && $('masterICTable').querySelector('tbody');
    if (icTb) {
      icTb.innerHTML = MASTER.icMembers
        .map(
          (ic) => `
      <tr>
        <td>${escapeHtml(ic)}</td>
        <td>${escapeHtml(MASTER.icToZone[ic] || '—')}</td>
        <td>${escapeHtml(MASTER.icToRic[ic] || '—')}</td>
        <td class="mono">${(MASTER.icToPartners[ic] || []).length}</td>
      </tr>`,
        )
        .join('');
    }
    const adminTb = $('adminTable') && $('adminTable').querySelector('tbody');
    if (adminTb) {
      // Show the canonical 4 authorised admins (from CONFIG.ADMIN_ALLOWLIST)
      adminTb.innerHTML = CONFIG.ADMIN_ALLOWLIST.map(
        (a) => `<tr><td>${escapeHtml(a)}</td></tr>`,
      ).join('');
    }
  }

  /* =========================================================
               22. PAGE ROUTER
            ========================================================= */
  function go(page) {
    // Role guards
    if (page === 'master' && state.role !== 'ADMIN') page = 'dashboard';
    if (page === 'entry' && state.role === 'VIEWER') page = 'dashboard';
    state.page = page;
    document
      .querySelectorAll('.page')
      .forEach((p) => p.classList.remove('active'));
    document
      .querySelectorAll('.nav-tab')
      .forEach((t) => t.classList.remove('active'));
    const target = $('page-' + page);
    if (target) target.classList.add('active');
    const tab = document.querySelector('.nav-tab[data-page="' + page + '"]');
    if (tab) tab.classList.add('active');
    if (page === 'dashboard') renderDashboard();
    if (page === 'records') renderRecords();
    if (page === 'master') renderMasterPage();
  }

  function refreshAll() {
    // ALWAYS recompute role-filtered records
    // after any role/user change.
    const visibleRecords = getVisibleRecords();

    // Header labels
    if ($('currentUserLabel')) {
      $('currentUserLabel').textContent = state.user || '—';
    }

    if ($('currentRoleLabel')) {
      $('currentRoleLabel').textContent = state.role || 'IC_MEMBER';
    }

    // Dashboard
    renderDashboard();

    // Records table
    renderRecords();

    // Reports
    // renderReports();

    // Audit logs
    // renderAudit();

    // Master page
    renderMasterPage();

    // Refresh report dropdowns
    refreshReportUserSelect();

    // Apply permissions/UI rules
    applyRolePermissions();

    // Update nav count badge
    const badge = document.querySelector(
      '.nav-tab[data-page="records"] .count',
    );

    if (badge) {
      badge.textContent = visibleRecords.length;
    }

    // Sync role dropdown enable/disable
    syncRoleDropdownState();
  }

  /* =========================================================
               23. BOOT
            ========================================================= */
  function init() {
    loadMaster();
    loadMasterFromGoogleSheet();
    populateMaster();
    refreshUserSelect();

    // Default IC meeting date = today
    $('f_date').value = new Date().toISOString().slice(0, 10);

    // Restore role + user from storage
    $('currentRole').value = state.role;
    // Re-validate admin authorization on boot
    if (state.role === 'ADMIN' && !isAdminAuthorized(state.user)) {
      state.role = 'IC_MEMBER';
      saveJSON(KEY_ROLE, state.role);
      $('currentRole').value = 'IC_MEMBER';
    }
    refreshUserSelect();
    syncRoleDropdownState();
    if (state.user) $('currentUser').value = state.user;

    // Restore draft
    loadDraft();

    // Session clock
    setInterval(() => {
      $('sessionTime').textContent = new Date().toLocaleTimeString('en-IN', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      });
    }, 1000);

    applyRolePermissions();
    startDraftAutoSave();

    // Inject Export/Import Data buttons into header
    const hm = document.querySelector('.header-meta');
    if (hm) {
      const persistDiv = document.createElement('div');
      persistDiv.className = 'item';
      persistDiv.style.cssText = 'display:flex;gap:6px;align-items:center;';
      persistDiv.innerHTML = `
      <button onclick="App.exportDataJSON()" class="btn btn-gold btn-sm" title="Export all records as JSON backup" style="font-size:10px;padding:5px 9px;">⬇ Export Data</button>
      <button onclick="App.importDataJSON()" class="btn btn-ghost btn-sm" title="Import a JSON backup to merge records" style="font-size:10px;padding:5px 9px;border-color:rgba(201,169,97,.5);color:var(--gold-500);">⬆ Import Data</button>`;
      hm.insertBefore(persistDiv, hm.firstChild);
    }

    // Initialise real-time sync (Firebase + BroadcastChannel + storage events)
    Sync.init();

    renderMasterPage();
    refreshAll();
  }

  /* Public API */
  return {
    go,
    submitEntry,
    saveDraft,
    clearForm,
    onMemberChange,
    renderRecords,
    resetFilters,
    showDetail,
    closeModal,
    adminDelete,
    editEntry,
    cancelEdit,
    exportAll,
    exportFiltered,
    exportOne,
    exportByDate,
    exportByZone,
    exportByPartner,
    exportByUser,
    exportMISPack,
    exportAudit,
    emailReadySummary,
    exportDataJSON,
    importDataJSON,
    handleMasterUpload,
    exportCurrentMaster,
    resetMaster,
    saveFirebaseURL,
    clearFirebaseURL,
    copyShareLink,
    setUser,
    setRole,
    openAdminPasswordModal,
    closeAdminPasswordModal,
    verifyAdminPassword,
    toggleAdminPassword,
    toggleICPassword,
    openICPasswordModal,
    closeICPasswordModal,
    verifyICPassword,
    toast,
    _init: init,
    _state: () => state,
    _master: () => MASTER,
  };
})();

window.addEventListener('DOMContentLoaded', App._init);
