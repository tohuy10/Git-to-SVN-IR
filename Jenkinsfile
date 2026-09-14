pipeline {
    agent any

    options {
        timestamps()
    }

    triggers {
        // Runs every Monday morning between 02:00 and 03:00 GMT+7
        cron('H 2 * * 1')
    }

    parameters {
        string(
            name: 'REPO_LIST',
            defaultValue: '',
            description: 'Comma-separated list of GitHub repository URLs (e.g. https://github.com/owner/repo1,https://github.com/owner/repo2)'
        )
        password(
            name: 'GITHUB_TOKEN',
            defaultValue: '',
            description: 'GitHub Personal Access Token (Masked)'
        )
        string(
            name: 'SVN_URL',
            defaultValue: '',
            description: 'Target SVN directory URL (e.g. svn://172.19.0.1:3690/company_repo/Project_Engineering/6_Review/Git_review_log/)'
        )
        string(
            name: 'SVN_USER',
            defaultValue: '',
            description: 'SVN Username'
        )
        password(
            name: 'SVN_PASS',
            defaultValue: '',
            description: 'SVN Password (Masked)'
        )
    }

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
                    # Export parameters directly into the shell session
                    export REPO_LIST="${REPO_LIST}"
                    export GITHUB_TOKEN="${GITHUB_TOKEN}"
                    export SVN_URL="${SVN_URL}"
                    export SVN_USER="${SVN_USER}"
                    export SVN_PASS="${SVN_PASS}"

                    python3 --version
                    svn --version | head -n 1

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
