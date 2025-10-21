from datetime import datetime


def greet(name: str) -> str:
    return f"שלום {name}! השעה עכשיו: {datetime.now():%H:%M:%S}"

if __name__ == "__main__":
    print(greet("אלעד"))
