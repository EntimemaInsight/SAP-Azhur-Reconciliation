import argparse,json
from .engine import run

def main():
    parser=argparse.ArgumentParser(description="Run SAP–Azhur supplier reconciliation")
    parser.add_argument("--config",default="config/settings.json")
    args=parser.parse_args();print(json.dumps(run(args.config),ensure_ascii=False,indent=2,default=str))
if __name__=="__main__":main()
