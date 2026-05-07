# AWS Cloud Computing — Infrastructure as Code Lab Project

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![IaC: Terraform](https://img.shields.io/badge/IaC-Terraform-7B42BC?logo=terraform&logoColor=white)](https://www.terraform.io/)
[![Cloud: AWS](https://img.shields.io/badge/Cloud-AWS-FF9900?logo=amazon-aws&logoColor=white)](https://aws.amazon.com/)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Region: ap-south-1](https://img.shields.io/badge/Region-ap--south--1-232F3E?logo=amazon-aws&logoColor=white)](https://aws.amazon.com/about-aws/global-infrastructure/regions_az/)

A complete, reproducible cloud-engineering project that provisions and
documents nine distinct AWS workloads using **Terraform** and **AWS
CloudFormation**. Every workload is defined declaratively, deployed end to
end, evidenced by AWS Console screenshots, and compiled into a single
formatted lab report.

The work was carried out for the *Cloud Computing* practical of the **B.E.
Computer Engineering** programme at **Far Western University, Faculty of
Engineering**, but is published here as a standalone, portfolio-grade
reference for any reader interested in hands-on AWS infrastructure
automation.

---

## Table of Contents

- [What this project demonstrates](#what-this-project-demonstrates)
- [Workloads](#workloads)
- [Architecture & approach](#architecture--approach)
- [Repository layout](#repository-layout)
- [Prerequisites](#prerequisites)
- [Reproducing a workload](#reproducing-a-workload)
- [Personalize this template](#personalize-this-template)
- [Rebuilding the lab report](#rebuilding-the-lab-report)
- [Cost discipline](#cost-discipline)
- [Process documentation](#process-documentation)
- [License](#license)

---

## What this project demonstrates

- **Infrastructure as Code at depth** — eight Terraform modules and one
  CloudFormation template, written from scratch. No copy-paste from blog
  posts, no premade modules. Every resource, every IAM policy, every wiring
  between services is explicit.
- **Breadth across AWS services** — networking (VPC, subnets, route tables,
  IGW, security groups), compute (EC2, ASG, ALB), storage (S3 with static
  hosting), identity (IAM users, groups, managed and inline policies),
  serverless (Lambda, DynamoDB, ECS Fargate), messaging (SNS pub/sub with
  SQS and email fan-out) and orchestration (CloudFormation stack lifecycle).
- **Engineering process** — the project was scoped via a written design
  spec, broken down into a step-by-step implementation plan, executed phase
  by phase, and version-controlled with a clean linear commit history. The
  full process documentation lives under [`docs/superpowers/`](./docs/superpowers).
- **Reproducibility** — anyone with an AWS account can clone the repo, run
  `terraform apply` in any lab folder, and recreate the exact environment
  shown in the screenshots within seconds.

---

## Workloads

| # | Workload | IaC | Key AWS services |
|---|---------|-----|------------------|
| 1 | Virtual Cloud Environment (VPC)               | Terraform      | VPC |
| 2 | Compute Instances with Startup Scripts        | Terraform      | EC2, AMI lookup, Security Groups |
| 3 | Object Storage with Static Website Hosting    | Terraform      | S3 (website endpoint, bucket policy) |
| 4 | Multi-tier Virtual Networking                 | Terraform      | VPC, Subnets, Route Tables, IGW, SGs |
| 5 | Load Balancer with Auto-Scaling Group         | Terraform      | ALB, Target Group, Launch Template, ASG, EC2 |
| 6 | Identity & Access Management                  | Terraform      | IAM Users, Groups, Managed Policies, Login Profile |
| 7 | Serverless Stack (FaaS + NoSQL + Containers)  | Terraform      | Lambda (Python 3.12), DynamoDB, ECS Fargate, IAM, CloudWatch |
| 8 | Pub/Sub Messaging with Fan-out                | Terraform      | SNS, SQS, SQS access policy, Email subscription |
| 9 | Declarative Infrastructure                    | CloudFormation | CloudFormation stack, VPC, S3 |

Each workload folder contains the IaC source plus a `README-screenshots.md`
that lists exactly which AWS Console pages were captured for evidence.

---

## Architecture & approach

Each workload is fully isolated — its own folder under
[`terraform/`](./terraform), its own state, its own resources tagged with
`Owner = Mukesh`, `Project = FWU-CloudComputing-Lab`, and `Lab = lab-NN`.
This isolation guarantees that running or destroying one workload never
disturbs another, which makes the project easy to learn from one piece at a
time.

Two design choices worth calling out:

- **Hybrid IaC.** Terraform was chosen for eight of the nine workloads because
  of its conciseness and readable HCL syntax. CloudFormation was used for the
  ninth specifically to demonstrate AWS-native declarative provisioning, since
  the syllabus topic itself is "Infrastructure as Code". The two approaches
  are compared in the final lab's *Observations* section.
- **Hybrid documentation.** Resources are provisioned via IaC, but the lab
  report describes the workflow in console-driven prose so that a reader who
  has never touched Terraform can follow the same set of clicks in the AWS
  Console and arrive at the same result.

---

## Repository layout

```
terraform/                          One folder per workload (Lab 9 uses CloudFormation YAML)
├── lab-01-vpc/
├── lab-02-ec2/
├── lab-03-s3-static-website/
├── lab-04-vpc-networking/
├── lab-05-alb-asg/
├── lab-06-iam/
├── lab-07-lambda-fargate-dynamodb/
├── lab-08-sns-sqs/
└── lab-09-cloudformation/

screenshots/                        AWS Console evidence per workload
scripts/build_report.py             python-docx report assembler
report/                             Compiled .docx lab report
docs/superpowers/                   Design spec + implementation plan
```

Each Terraform folder follows the same conventions: `provider.tf` (region
and default tags), `variables.tf`, `main.tf`, `outputs.tf`, plus
workload-specific files where needed (`user_data.sh`, `lambda/handler.py`,
`web/index.html`, etc.).

---

## Prerequisites

| Tool | Minimum version | Purpose |
|------|-----------------|---------|
| AWS account | Free Tier or paid | Target environment |
| Terraform   | 1.6+            | Eight of the nine workloads |
| AWS CLI     | v2              | CloudFormation deploy + general account access |
| Python      | 3.11+           | Rebuild the .docx report (optional) |

The default region used throughout is `ap-south-1` (Mumbai). Override in
`provider.tf` if you prefer a different region.

---

## Reproducing a workload

Pick any folder under `terraform/`:

```bash
cd terraform/lab-01-vpc
terraform init
terraform apply -auto-approve

# Capture the screenshots listed in README-screenshots.md, then:

terraform destroy -auto-approve
```

A few workloads have specifics worth noting:

- **Lab 8** requires an email at apply time so SNS can deliver:
  `terraform apply -auto-approve -var="subscriber_email=you@example.com"` —
  remember to confirm the AWS subscription email in your inbox.
- **Lab 9** is CloudFormation-driven:
  `aws cloudformation deploy --template-file stack.yaml --stack-name stack-mukesh --region ap-south-1 --capabilities CAPABILITY_NAMED_IAM`.
- **Labs 5 and 7** provision resources outside the strict AWS Free Tier
  (an Application Load Balancer, an ECS Fargate task). End-to-end run cost
  is typically **under USD 2** when destroyed promptly.

---

## Personalize this template

The repository is published as a personalizable template. Resource names,
tags, the rendered HTML pages, the Lambda payload, the CloudFormation
parameter, and every line of the compiled lab report all read from a
single root file: [`config.yaml`](./config.yaml). Out of the box, the
defaults reproduce the author's own resources and report; with one edit,
the entire project re-skins to a different identity.

### Option 1 — Interactive wizard

```bash
python scripts/personalize.py
```

The wizard prompts for every value (resource suffix, display name, full
name, roll number, semester, program, university, professor, AWS region).
Press Enter on any prompt to keep the current default. After the prompts
it propagates the values into per-lab `terraform.auto.tfvars.json`
(auto-loaded by Terraform) and rewrites the default `OwnerName`
parameter in Lab 9's CloudFormation template.

### Option 2 — Edit `config.yaml` directly

Open [`config.yaml`](./config.yaml), change the values you care about,
save, then run a non-interactive propagation:

```bash
python scripts/personalize.py --apply --no-prompt
```

### Option 3 — One-shot CLI flags

```bash
python scripts/personalize.py \
  --suffix jane \
  --display-name Jane \
  --full-name "Jane Doe" \
  --roll 5 \
  --semester VII \
  --professor "Dr. Alice Smith" \
  --apply --no-prompt
```

### After personalizing

- `terraform apply -auto-approve` in any lab folder uses *your* values —
  resources will be named `vpc-jane`, `ec2-jane`, and so on.
- `python scripts/build_report.py` produces
  `report/<Your_Name>_Cloud_Computing_Lab_Report.docx` with a cover page,
  figure captions, and procedure prose all referencing your identity.
- The screenshots committed under `screenshots/` are the original
  author's. Capture your own per each lab's `README-screenshots.md` and
  drop them into the same paths to overwrite.

The `subject_email` for Lab 8 is intentionally **not** in `config.yaml` —
it's a runtime CLI argument
(`-var="subscriber_email=you@example.com"`) so personal email addresses
never get committed to the repo.

---

## Rebuilding the lab report

The compiled report at
[`report/Mukesh_Pant_Cloud_Computing_Lab_Report.docx`](./report) is
generated programmatically from the screenshots and prose contained in
`scripts/build_report.py`.

```bash
pip install python-docx Pillow
python scripts/build_report.py
```

Editing the `LAB_*` dictionaries at the top of the script and re-running
regenerates the entire document with new content, layout, or styling
without ever touching Word manually.

---

## Cost discipline

- Run `terraform destroy` after each workload — the IaC layer makes this
  trivial.
- All EC2 instances use `t2.micro` (Free Tier eligible in `ap-south-1`).
- DynamoDB uses on-demand billing — minimal usage in this project incurs
  effectively zero cost.
- The ALB (Lab 5) has an hourly fee while running; destroying it within 30
  minutes keeps the bill near zero.

---

## Process documentation

Two artefacts under [`docs/superpowers/`](./docs/superpowers) capture how
the project was scoped before any code was written:

- [`specs/2026-04-27-cloud-lab-report-design.md`](./docs/superpowers/specs/2026-04-27-cloud-lab-report-design.md)
  — design spec covering goals, scope, risks, free-tier safety, and an
  outline of every workload.
- [`plans/2026-04-27-cloud-lab-report-implementation.md`](./docs/superpowers/plans/2026-04-27-cloud-lab-report-implementation.md)
  — step-by-step implementation plan: which files to create, which AWS
  resources to provision, which screenshots to capture, and how to clean up.

Together they show the project moving from problem statement to working
infrastructure to compiled report in a deliberate, reviewable sequence
rather than as ad-hoc clicking around the AWS Console.

---

## License

Released under the [MIT License](./LICENSE) — © 2026 Mukesh Pant.
