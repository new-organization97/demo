import argparse
import requests
import sys
import json
import os
import dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from datetime import datetime
from typing import List, Optional

dotenv.load_dotenv()

github_token = os.getenv("TOKEN")
if not github_token:
    print("GITHUB_TOKEN environment variable is not set.")
    sys.exit(1)

# --- Excel Logging Configuration ---
# Assuming the script is run from the root of the repo
EXCEL_FILE_PATH = os.path.join("logs", "github_admin_log.xlsx")

def log_action_to_excel(action_details: dict):
    """
    Appends action details to the Excel log file.
    """
    try:
        # Load the workbook or create a new one if it doesn't exist
        if os.path.exists(EXCEL_FILE_PATH):
            workbook = load_workbook(EXCEL_FILE_PATH)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            # Add headers if it's a new file
            sheet.append([
                "Timestamp (IST)", "Action", "Organization", "Team", "Repository",
                "User", "Permission", "New Repo Name", "Private Repo"
            ])
            # Apply bold style to headers (optional, but good for readability)
            for cell in sheet[1]: # Iterate through cells in the first row
                cell.font = Font(bold=True)
            print(f"Created new Excel log file with headers: {EXCEL_FILE_PATH}")

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
        # This will be boolean, convert to string for excel. Ensure it's handled properly for logging.
        # If it's a boolean False, str(False) is "False", which is fine.
        repo_private = action_details.get("repo_private", False) 

        sheet.append([
            timestamp, action, org, team, repo,
            user, permission, repo_name, str(repo_private)
        ])

        # Save the workbook
        workbook.save(EXCEL_FILE_PATH)
        print(f"✅ Action logged successfully to {EXCEL_FILE_PATH}")
    except Exception as e:
        print(f"❌ Error logging action to Excel: {e}")
        # Consider re-raising for critical errors or adding more robust error handling

# --- End Excel Logging Configuration ---


class GitHubAPIManager:
    def __init__(self, token: str):
        self.token = token
        self.base_url = "https://api.github.com"
        self.headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "GitHub-Manager-Script"
        }

    def make_request(self, method: str, endpoint: str, data: dict = None) -> dict:
        """Make HTTP request to GitHub API"""
        url = f"{self.base_url}{endpoint}"
        
        try:
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
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Request failed: {str(e)}")
            return None
        except json.JSONDecodeError:
            print(f"❌ Invalid JSON response")
            return None

    def list_orgs(self) -> List[str]:
        """List organizations user is a member of"""
        response = self.make_request("GET", "/user/memberships/orgs")
        if response is None:
            return []
        return [org['organization']['login'] for org in response]

    def list_teams(self, org: str) -> List[dict]:
        """List teams in an organization"""
        response = self.make_request("GET", f"/orgs/{org}/teams")
        if response is None:
            return []
        return response

    def list_repos(self, org: str) -> List[dict]:
        """List repositories in an organization"""
        response = self.make_request("GET", f"/orgs/{org}/repos")
        if response is None:
            return []
        return response

    def create_team(self, org: str, team_name: str, description: str = "") -> bool:
        """Create a team in an organization"""
        data = {
            "name": team_name,
            "description": description,
            "privacy": "closed"  
        }
        response = self.make_request("POST", f"/orgs/{org}/teams", data)
        if response:
            print(f"✅ Created team '{team_name}' in '{org}'")
            return True
        return False

    def delete_team(self, org: str, team_slug: str) -> bool:
        """Delete a team from an organization"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}")
        if response is not None:
            print(f"✅ Deleted team '{team_slug}' in '{org}'") # Changed ❌ to ✅
            return True
        return False

    def add_team_to_repo(self, org: str, team_slug: str, repo: str, permission: str) -> bool:
        """Add team to repository with specific permission"""
        data = {"permission": permission}
        response = self.make_request("PUT", f"/orgs/{org}/teams/{team_slug}/repos/{org}/{repo}", data)
        if response is not None:
            print(f"✅ Added team '{team_slug}' to repo '{repo}' with permission '{permission}'")
            return True
        return False

    def remove_team_from_repo(self, org: str, team_slug: str, repo: str) -> bool:
        """Remove team from repository"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}/repos/{org}/{repo}")
        if response is not None:
            print(f"✅ Removed team '{team_slug}' from repo '{repo}'") # Changed ❌ to ✅
            return True
        return False

    def add_user_to_team(self, org: str, team_slug: str, username: str) -> bool:
        """Add user to team"""
        response = self.make_request("PUT", f"/orgs/{org}/teams/{team_slug}/memberships/{username}")
        if response:
            print(f"✅ Added user '{username}' to team '{team_slug}' in '{org}'")
            return True
        return False

    def remove_user_from_team(self, org: str, team_slug: str, username: str) -> bool:
        """Remove user from team"""
        response = self.make_request("DELETE", f"/orgs/{org}/teams/{team_slug}/memberships/{username}")
        if response is not None:
            print(f"✅ Removed user '{username}' from team '{team_slug}' in '{org}'") # Changed ❌ to ✅
            return True
        return False

    def create_repo(self, org: str, repo_name: str, private: bool = False, description: str = "") -> bool:
        """Create repository in organization"""
        data = {
            "name": repo_name,
            "description": description,
            "private": private,
            "has_issues": True,
            "has_projects": True,
            "has_wiki": True
        }
        response = self.make_request("POST", f"/orgs/{org}/repos", data)
        if response:
            visibility = "private" if private else "public"
            print(f"✅ Created {visibility} repo '{repo_name}' in '{org}'")
            return True
        return False

    def validate_user(self, username: str) -> bool:
        """Validate if GitHub user exists"""
        if "@" in username:
            print(f"❌ Email detected: '{username}' — GitHub API requires the GitHub username instead.")
            print("👉 Please enter the GitHub username (e.g. 'pirai-santhosh'), not the email address.")
            return False

        response = self.make_request("GET", f"/users/{username}")
        return response is not None

    def get_user_repo_access(self, org: str, username: str) -> List[str]:
        """Get list of repositories user has access to in organization"""
        repos = self.list_repos(org)
        access_repos = []

        for repo in repos:
            # Check if user is a collaborator
            response = self.make_request("GET", f"/repos/{org}/{repo['name']}/collaborators/{username}")
            if response is not None:
                access_repos.append(repo['name'])

        print(f"📆 User '{username}' has access to {len(access_repos)} repositories in '{org}':")
        for repo in access_repos:
            print(f"  - {repo}")
        
        return access_repos

    def get_team_by_name(self, org: str, team_name: str) -> Optional[dict]:
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
        "repo_private": args.repo_private if hasattr(args, 'repo_private') else False # Default to False if not set
    }

    if args.action == "list-orgs":
        orgs = github.list_orgs()
        print("📋 Organizations:")
        for org in orgs:
            print(f"  - {org}")
        # Not typically logged to Excel as it's just a read operation
    
    elif args.action == "list-teams":
        teams = github.list_teams(args.org)
        print(f"📋 Teams in organization '{args.org}':")
        if teams:
            for i, team in enumerate(teams, 1):
                print(f"  {i}. {team['name']} (ID: {team['id']}, Slug: {team['slug']})")
        else:
            print("  No teams found.")
        # Not typically logged to Excel as it's just a read operation

    elif args.action == "list-repos":
        repos = github.list_repos(args.org)
        print(f"📋 Repositories in organization '{args.org}':")
        if repos:
            for i, repo in enumerate(repos, 1):
                visibility = "🔒 Private" if repo['private'] else "🌐 Public"
                print(f"  {i}. {repo['name']} ({visibility})")
        else:
            print("  No repositories found.")
        # Not typically logged to Excel as it's just a read operation

    elif args.action == "create-team":
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
            # Corrected error message here
            print("--team, --repo, and --permission are required for add-repo")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            if github.add_team_to_repo(args.org, team_info['slug'], args.repo, args.permission):
                log_data["team_slug"] = team_info['slug'] # Add slug for logging if needed
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
        
        log_data["team_slug"] = team_info['slug'] # Add slug for logging if needed

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

    elif args.action == "user-access":
        if not args.user:
            print("--user is required for user-access")
            sys.exit(1)
        if github.validate_user(args.user): # Validate user before proceeding
            github.get_user_repo_access(args.org, args.user)
        # Not typically logged to Excel as it's just a read operation,
        # but you could if you wanted to track when reports are generated.

def main():
    parser = argparse.ArgumentParser(description="GitHub Team and Repo Manager (Direct API)")
    
    parser.add_argument("--action",
                        choices=[
                            "create-team", "delete-team", "add-repo", "remove-repo",
                            "add-user", "remove-user", "create-repo", "user-access",
                            "list-teams", "list-repos", "list-orgs" # Added list actions
                        ],
                        required=True,
                        help="Action to perform")
    
    parser.add_argument("--org", help="GitHub organization name")
    parser.add_argument("--team", help="Team name")
    parser.add_argument("--repo", help="Repository name")
    parser.add_argument("--user", help="GitHub username (not email!)")
    parser.add_argument("--permission",
                        choices=["pull", "triage", "push", "maintain", "admin"],
                        help="Permission level for team access to repository")
    parser.add_argument("--repo-private", action="store_true",
                        help="Create repository as private (default is public)")
    parser.add_argument("--repo-name", help="Name for new repository")

    args = parser.parse_args()
    
    # Check if org is required for the action
    if args.action not in ["list-orgs"] and not args.org:
        print(f"--org is required for {args.action}")
        sys.exit(1)
    
    run_action(args)

if __name__ == "__main__":
    main()
