pipeline {
    agent any

    // options {
    //     timestamps()
    // }

    // NOTE: Comment out 'triggers' in script to avoid wiping out the 
    // UI-configured schedule.
    // triggers {
    //     // Runs every Monday morning between 02:00 and 03:00 GMT+7
    //     cron('H 2 * * 1')
    // }
    
    // NOTE: 'parameters' block commented out so Jenkins does NOT 
    // wipe the UI-configured default passwords on weekly runs.
    // parameters {
    //     string(
    //         name: 'REPO_LIST',
    //         defaultValue: '',
    //         description: 'Comma-separated list of GitHub repository URLs (e.g. https://github.com/owner/repo1,https://github.com/owner/repo2)'
    //     )
    //     password(
    //         name: 'GITHUB_TOKEN',
    //         defaultValue: '',
    //         description: 'GitHub Personal Access Token (Masked)'
    //     )
    //     string(
    //         name: 'SVN_URL',
    //         defaultValue: '',
    //         description: 'Target SVN directory URL (e.g. svn://172.19.0.1:3690/company_repo/Project_Engineering/6_Review/Git_review_log/)'
    //     )
    //     string(
    //         name: 'SVN_USER',
    //         defaultValue: '',
    //         description: 'SVN Username'
    //     )
    //     password(
    //         name: 'SVN_PASS',
    //         defaultValue: '',
    //         description: 'SVN Password (Masked)'
    //     )
    // }

    stages {
        stage('1. Validate Inputs') {
            steps {
                script {
                    if (!params.REPO_LIST?.trim()) {
                        error("Parameter REPO_LIST must not be empty.")
                    }
                    if (!params.GITHUB_TOKEN?.toString()?.trim()) {
                        error("Parameter GITHUB_TOKEN must not be empty.")
                    }

                    if (!params.SVN_URL?.trim()) {
                        error("Parameter SVN_URL must not be empty.")
                    }
                }
            }
        }

        stage('2. Checkout Script Repository') {
            steps {
                git branch: 'main', url: 'https://github.com/tohuy10/Git-to-SVN-IR.git'
            }
        }

        stage('3. Execute Review Sync') {
            steps {
                sh '''
                    export PATH="$HOME/.local/bin:$PATH"

                    # 1. Install pip if not present
                    if ! python3 -m pip --version >/dev/null 2>&1; then
                        echo "[*] Pip not found, bootstrapping..."
                        curl -sS https://bootstrap.pypa.io/get-pip.py -o get-pip.py || wget -q https://bootstrap.pypa.io/get-pip.py -O get-pip.py
                        python3 get-pip.py --user --no-warn-script-location
                        rm -f get-pip.py
                    fi

                    # 2. Install required Python packages
                    python3 -m pip install --user --upgrade requests openpyxl

                    # 3. Print environment info
                    python3 --version
                    svn --version | head -n 1 || true

                    # 4. Run the review sync script
                    python3 sync_reviews.py
                '''
            }
        }
    }

    post {
        always {
            // Archive the generated Excel log so you can download it directly from the Jenkins UI
            archiveArtifacts artifacts: 'svn_workdir/*.xlsx', allowEmptyArchive: true
        }
        success {
            echo "Review logs successfully synced to SVN."
        }
        failure {
            echo "Review sync job failed. Check console output for details."
        }
    }
}
