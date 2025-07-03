import sys
import os

if __name__ == "__main__":
    # Set debug environment variable
    os.environ['DEBUG'] = '1'
    
    # Add src directory to Python path
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
    
    # Import and run main
    from main import main
    main()