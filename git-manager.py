import argparse
import requests
import sys
import json
import os
from openpyxl import Workbook, load_workbook  
from datetime import datetime


github_token = os.getenv("TOKEN")
if not github_token:
    print("GITHUB_TOKEN environment variable is not set.")
    sys.exit(1)

# Excel Logging Configuration
EXCEL_FILE_PATH = "logs/github_admin_log.xlsx"

def log_action_to_excel(action_details: dict):
    if os.path.exists(EXCEL_FILE_PATH):
        workbook = load_workbook(EXCEL_FILE_PATH)
        sheet = workbook.active
    else:
        print(f"❌ Error: Excel log file not found at {EXCEL_FILE_PATH}. Script will exit.")
        return

    # Append the new row of data
    # Ensure all possible keys are present, even if empty, to maintain column consistency
    timestamp = action_details.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    action = action_details.get("action", "")
    org = action_details.get("org", "")
    team = action_details.get("team", "")
    repo = action_details.get("repo", "")
    user = action_details.get("user", "")
    permission = action_details.get("permission", "")
    repo_name = action_details.get("repo_name", "")
    repo_private = action_details.get("repo_private", "") # This will be boolean, convert to string for excel

    sheet.append([
        timestamp, action, org, team, repo,
        user, permission, repo_name, str(repo_private)
    ])

    # Save the workbook
    workbook.save(EXCEL_FILE_PATH)
    print(f"✅ Action logged successfully to {EXCEL_FILE_PATH}")      

class GitHubAPIManager:
    def __init__(self, token: str):
        self.token = token
        self.base_url = "https://api.github.com"
        self.headers = {
            "Authorization": f"token {self.token}"
        }

    def make_request(self, method: str, endpoint: str, data: dict = None):
        """Make HTTP request to GitHub API"""
        url = f"{self.base_url}{endpoint}"
        
        if method.upper() == "GET":
            response = requests.get(url, headers=self.headers)
        elif method.upper() == "POST":
            response = requests.post(url, headers=self.headers, json=data)
        elif method.upper() == "PUT":
            response = requests.put(url, headers=self.headers, json=data)
        elif method.upper() == "DELETE":
            response = requests.delete(url, headers=self.headers)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
            
        if response.status_code not in [200, 201, 204]:
            error_msg = response.json().get('message', 'Unknown error') if response.text else 'No response'
            print(f"❌ API Error ({response.status_code}): {error_msg}")
            return None
            
        return response.json() if response.text else {}

    def list_teams(self, org: str):
        response = self.make_request("GET", f"/orgs/{org}/teams")
        if response is None:
            return []
        return response

    def create_team(self, org: str, team_name: str):
        """Create a team in an organization"""
        data = {
            "name": team_name,
            "privacy": "closed"  
        }
        response = self.make_request("POST", f"/orgs/{org}/teams", data)
        if response:
            print(f"✅ Created team '{team_name}' in '{org}'")
            return True
        return False

    def delete_team(self, org: str, team_slug: str):
        """Delete a team from an organization"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}")
        if response is not None:
            print(f"❌ Deleted team '{team_slug}' in '{org}'")
            return True
        return False

    def add_team_to_repo(self, org: str, team_slug: str, repo: str, permission: str):
        """Add team to repository with specific permission"""
        data = {"permission": permission}
        response = self.make_request("PUT", f"/orgs/{org}/teams/{team_slug}/repos/{org}/{repo}", data)
        if response is not None:
            print(f"✅ Added team '{team_slug}' to repo '{repo}' with permission '{permission}'")
            return True
        return False

    def remove_team_from_repo(self, org: str, team_slug: str, repo: str):
        """Remove team from repository"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}/repos/{org}/{repo}")
        if response is not None:
            print(f"❌ Removed team '{team_slug}' from repo '{repo}'")
            return True
        return False

    def add_user_to_team(self, org: str, team_slug: str, username: str):
        """Add user to team"""
        response = self.make_request("PUT", f"/orgs/{org}/teams/{team_slug}/memberships/{username}")
        if response:
            print(f"✅ Added user '{username}' to team '{team_slug}' in '{org}'")
            return True
        return False

    def remove_user_from_team(self, org: str, team_slug: str, username: str):
        """Remove user from team"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}/memberships/{username}")
        if response is not None:
            print(f"❌ Removed user '{username}' from team '{team_slug}' in '{org}'")
            return True
        return False

    def create_repo(self, org: str, repo_name: str, private: bool = False):
        """Create repository in organization"""
        data = {
            "name": repo_name,
            "private": private,
        }
        response = self.make_request("POST", f"/orgs/{org}/repos", data)
        if response:
            visibility = "private" if private else "public"
            print(f"✅ Created {visibility} repo '{repo_name}' in '{org}'")
            return True
        return False

    def validate_user(self, username: str):
        """Validate if GitHub user exists"""
        if "@" in username:
            print(f"❌ Email detected: '{username}' — GitHub API requires the GitHub username instead.")
            print("👉 Please enter the GitHub username (e.g. 'pirai-deepak'), not the email address.")
            return False

        response = self.make_request("GET", f"/users/{username}")
        return response is not None

    def get_team_by_name(self, org: str, team_name: str):
        """Find team by name and return team info"""
        teams = self.list_teams(org)
        for team in teams:
            if team['name'].lower() == team_name.lower():
                return team
        return None

def run_action(args):
    github = GitHubAPIManager(github_token)
    
    # Get current timestamp for logging
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # The log_data dictionary will store details for Excel logging
    log_data = {
        "timestamp": current_time,
        "action": args.action,
        "org": args.org,
        "team": args.team if hasattr(args, 'team') else "",
        "repo": args.repo if hasattr(args, 'repo') else "",
        "user": args.user if hasattr(args, 'user') else "",
        "permission": args.permission if hasattr(args, 'permission') else "",
        "repo_name": args.repo_name if hasattr(args, 'repo_name') else "",
        "repo_private": args.repo_private if hasattr(args, 'repo_private') else False
    }

    if args.action == "create-team":
        if not args.team:
            print("--team is required for create-team")
            sys.exit(1)
        if github.create_team(args.org, args.team):
            log_action_to_excel(log_data)
    
    elif args.action == "delete-team":
        if not args.team:
            print("--team is required for delete-team")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            if github.delete_team(args.org, team_info['slug']):
                log_action_to_excel(log_data)
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
            sys.exit(1) # Exit if team not found for deletion

    elif args.action == "add-repo":
        if not all([args.team, args.repo, args.permission]):
            print("--team, --repo, and --permission are required for add-repo")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            if github.add_team_to_repo(args.org, team_info['slug'], args.repo, args.permission):
                log_action_to_excel(log_data)
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
            sys.exit(1) # Exit if team not found for adding repo

    elif args.action == "remove-repo":
        if not all([args.team, args.repo]):
            print("--team and --repo are required for remove-repo")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            if github.remove_team_from_repo(args.org, team_info['slug'], args.repo):
                log_action_to_excel(log_data)
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
            sys.exit(1) # Exit if team not found for removing repo

    elif args.action in ["add-user", "remove-user"]:
        if not all([args.team, args.user]):
            print("--team and --user are required for user management")
            sys.exit(1)

        if not github.validate_user(args.user):
            print(f"❌ Invalid GitHub username: {args.user}")
            sys.exit(1)

        team_info = github.get_team_by_name(args.org, args.team)
        if not team_info:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
            sys.exit(1)

        if args.action == "add-user":
            if github.add_user_to_team(args.org, team_info['slug'], args.user):
                log_action_to_excel(log_data)
        else: # remove-user
            if github.remove_user_from_team(args.org, team_info['slug'], args.user):
                log_action_to_excel(log_data)

    elif args.action == "create-repo":
        if not args.repo_name:
            print("--repo-name is required for create-repo")
            sys.exit(1)
        if github.create_repo(args.org, args.repo_name, args.repo_private):
            log_action_to_excel(log_data)

def main():
    parser = argparse.ArgumentParser()
    
    parser.add_argument("--action",
                        choices=[
                            "create-team", "delete-team", "add-repo", "remove-repo",
                            "add-user", "remove-user", "create-repo"
                        ],
                        required=True,)
    
    parser.add_argument("--org", required=True)
    parser.add_argument("--team")
    parser.add_argument("--repo")
    parser.add_argument("--user")
    parser.add_argument("--permission",
                        choices=["pull", "triage", "push", "maintain", "admin"])
    parser.add_argument("--repo-private", action="store_true")
    parser.add_argument("--repo-name")

    args = parser.parse_args()
    
    run_action(args)

main()
