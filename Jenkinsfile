pipeline {
    agent any

    parameters {
        string(
            name: 'REPO_LIST',
            defaultValue: '',
            description: 'Comma-separated list of GitHub repository URLs (e.g. https://github.com/owner/repo1,https://github.com/owner/repo2)'
        )
        string(
            name: 'SVN_URL',
            defaultValue: '',
            description: 'Target SVN directory URL (e.g. svn://172.19.0.1:3690/company_repo/Project_Engineering/6_Review/Git_review_log/)'
        )
    }

    environment {
        GITHUB_TOKEN = credentials('github-token')
        SVN_CREDS    = credentials('svn-credentials')
        SVN_USER     = "${SVN_CREDS_USR}"
        SVN_PASS     = "${SVN_CREDS_PSW}"

        REPO_LIST    = "${params.REPO_LIST}"
        SVN_URL      = "${params.SVN_URL}"
    }

    stages {
        stage('Validate Inputs') {
            steps {
                script {
                    if (!params.REPO_LIST?.trim()) {
                        error("Parameter REPO_LIST must not be empty.")
                    }
                    if (!params.SVN_URL?.trim()) {
                        error("Parameter SVN_URL must not be empty.")
                    }
                }
            }
        }

        stage('Checkout Script Repository') {
            steps {
                git branch: 'main', url: 'https://github.com/tohuy10/Git-to-SVN-IR.git'
            }
        }

        stage('Execute Review Sync') {
            steps {
                sh '''
                    python3 --version
                    svn --version | head -n 1

                    python3 sync_reviews.py
                '''
            }
        }
    }

    post {
        success {
            echo "Review logs successfully synced to SVN."
        }
        failure {
            echo "Review sync job failed. Check console output for details."
        }
    }
}