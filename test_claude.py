import os
from dotenv import load_dotenv

load_dotenv()


def test_claude_connection():
    try:
        import anthropic
    except ImportError:
        print("Anthropic SDK is not installed. Install it with: pip install anthropic")
        return

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ANTHROPIC_API_KEY is not set. Please set it in your environment.")
        return

    try:
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=100,
            messages=[
                {
                    "role": "user",
                    "content": "Say hello. I am testing my API connection."
                }
            ]
        )
        print("Claude API response:")
        print(message.content[0].text)
    except Exception as error:
        print("Failed to test Claude connection:", error)


def prompt_search():
    try:
        from sender import check_excel_and_match
    except Exception:
        print("Sender module unavailable or missing dependencies. Make sure sender.py and required libraries are present.")
        return

    excel_path = os.environ.get("EXCEL_PATH", "tickets.xlsx")

    while True:
        try:
            user_input = input(
                "Please enter a description of the issue. "
                "If you are searching by a number, specify whether it is an error code, ticket ID, "
                "case number, or status code (example: 'error: 401'). "
                "Type 'back' to return to menu: "
            ).strip()
        except (KeyboardInterrupt, EOFError):
            print("\nInput cancelled. Returning to menu.")
            return

        if not user_input:
            print("Please enter search criteria or type 'back' to return to menu.")
            continue

        normalized = user_input.lower().strip()
        if normalized in ("back", "menu"):
            return
        if normalized in ("exit", "quit"):
            print("Returning to menu.")
            return

        try:
            matches = check_excel_and_match(excel_path, user_input)
            print(f"check_excel_and_match ran just now")
            if not matches:
                print(f"No matches found in {excel_path} for: {user_input}")
            else:
                print(f"Found {len(matches)} match(es):")
                for m in matches:
                    print(m)
            # Continue prompting for additional searches
        except FileNotFoundError:
            print(f"Excel file not found: {excel_path}. Returning to menu.")
            return
        except ValueError as error:
            print(error)
            print("Please try again with more context.")
            continue
        except Exception as error:
            print("Search failed:", error)
            return


def main_menu():
    while True:
        print("\nSelect an option:")
        print("  test  - Check Claude API connection")
        print("  start - Search the Excel file")
        print("  exit  - Quit")

        try:
            choice = input("Enter option: ").strip().lower()
        except (KeyboardInterrupt, EOFError):
            print("\nExiting.")
            return

        if choice == "test":
            test_claude_connection()
        elif choice == "start":
            prompt_search()
        elif choice in ("exit", "quit"):
            print("Goodbye.")
            return
        else:
            print("Unknown option. Please enter 'test', 'start', or 'exit'.")


if __name__ == "__main__":
    main_menu()