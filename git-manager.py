import argparse
import requests
import sys
import json
import os
import dotenv
#from openpyxl import Workbook, load_workbook
#from openpyxl.styles import Font
from datetime import datetime
from typing import List, Optional
 
 
dotenv.load_dotenv()
 
github_token = os.getenv("TOKEN")
if not github_token:
    print("GITHUB_TOKEN environment variable is not set.")
    sys.exit(1)
 
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
            print(f"❌ Deleted team '{team_slug}' in '{org}'")
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
            print(f"❌ Removed team '{team_slug}' from repo '{repo}'")
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
            print(f"❌ Removed user '{username}' from team '{team_slug}' in '{org}'")
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
 
    if args.action == "list-orgs":
        orgs = github.list_orgs()
        print("📋 Organizations:")
        for org in orgs:
            print(f"  - {org}")
 
    elif args.action == "list-teams":
        teams = github.list_teams(args.org)
        print(f"📋 Teams in organization '{args.org}':")
        if teams:
            for i, team in enumerate(teams, 1):
                print(f"  {i}. {team['name']} (ID: {team['id']}, Slug: {team['slug']})")
        else:
            print("  No teams found.")
 
    elif args.action == "list-repos":
        repos = github.list_repos(args.org)
        print(f"📋 Repositories in organization '{args.org}':")
        if repos:
            for i, repo in enumerate(repos, 1):
                visibility = "🔒 Private" if repo['private'] else "🌐 Public"
                print(f"  {i}. {repo['name']} ({visibility})")
        else:
            print("  No repositories found.")
 
    elif args.action == "create-team":
        if not args.team:
            print("--team is required for create-team")
            sys.exit(1)
        github.create_team(args.org, args.team)
 
    elif args.action == "delete-team":
        if not args.team:
            print("--team is required for delete-team")
            sys.exit(1)
        # Convert team name to slug if needed
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            github.delete_team(args.org, team_info['slug'])
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
 
    elif args.action == "add-repo":
        if not all([args.team, args.repo, args.permission]):
            print("--team, --repo, and --permission are required for add-repo")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            github.add_team_to_repo(args.org, team_info['slug'], args.repo, args.permission)
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
 
    elif args.action == "remove-repo":
        if not all([args.team, args.repo]):
            print("--team and --repo are required for remove-repo")
            sys.exit(1)
        team_info = github.get_team_by_name(args.org, args.team)
        if team_info:
            github.remove_team_from_repo(args.org, team_info['slug'], args.repo)
        else:
            print(f"❌ Team '{args.team}' not found in '{args.org}'")
 
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
            github.add_user_to_team(args.org, team_info['slug'], args.user)
        else:
            github.remove_user_from_team(args.org, team_info['slug'], args.user)
 
    elif args.action == "create-repo":
        if not args.repo_name:
            print("--repo-name is required for create-repo")
            sys.exit(1)
        github.create_repo(args.org, args.repo_name, args.repo_private)
 
    elif args.action == "user-access":
        if not args.user:
            print("--user is required for user-access")
            sys.exit(1)
        github.get_user_repo_access(args.org, args.user)
 
def main():
    parser = argparse.ArgumentParser(description="GitHub Team and Repo Manager (Direct API)")
   
    parser.add_argument("--action",
                       choices=[
                           "create-team", "delete-team", "add-repo", "remove-repo",
                           "add-user", "remove-user", "create-repo", "user-access",
                           "list-teams", "list-repos", "list-orgs"
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
