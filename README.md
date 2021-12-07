# Usage

1. Clone openshift-doc and checkout your favorite branch.

```
git clone https://github.com/openshift/openshift-docs.git
cd ./openshift-docs/
git switch enterprise-4.8
cd ..
```

2. Clone this repo.

```
git clone https://github.com/orimanabu/openshift_rest_api_adoc2xlsx.git
```

3. Run the script.

```
cd ./openshift_rest_api_adoc2xlsx/
./do.sh ../openshift-docs ./xlsx
```

Then create `./xlsx` directory and convert all rest api refeference asciidocs in openshift-docs repo to xlsx files in `./xlsx` directory.

# Tips

When debugging the script, `-f json` option might be useful.

```
./adoc2xlsx.py ../openshift-docs/rest_api/workloads_apis/pod-core-v1.adoc -f json | jq .
```
