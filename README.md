acidoc
======

ACIDoc will document pertitent information about the configuration of a Cisco ACI fabric into Spreadsheet format,
it turns out sometimes that's still the best way to pass information around.

The contents and structure of the outputed format is very easy to customize
of the output is easily changed by modifying the [config.yml](config.yaml) file

# Installation

## Using PIP

    pip install -r requirements.txt

## Using Virtualenv

    virtualenv venv
    source venv/bin/activate
    pip install -r requirements.txt


 ## Methodology

Everything in ACI is an object, every object is instantiated from a class which defines properties of the object.  
Using the ACI REST API you can query for all objects (and their properties) of a given class.

To construct a spreadsheet, this utility creates a tab for each class query, where the columns
are properties which you are interested in, and each row is an instance of that class

## configuration


### What ACI fabric should we get the information from?

First we need to authenticate to the APIC properly, we use the [acitoolkit](https://github.com/datacenter/acitoolkit) library to help
with this.

The easiest way is to set up your credentials is by adding some environment variables. If you are a mac user it would look something like this.


```
my-macbook-prompt ✗ export APIC_URL=https://myapic
my-macbook-prompt ✗ export APIC_LOGIN=someuser
my-macbook-prompt ✗ export APIC_PASSWORD=somepass
```

If you don't want to set these, the utility will also prompt the user for the required connection information at runtime.. e.g.

    my-macbook ✗ python aci-doc.py
    APIC login username: admin
    APIC URL: https://10.94.140.72
    APIC Password:
    Connected to ACI
    
### What do we want to see?

Start by defining what you want the output spreadsheet to be named.

```
filename: aci-doc.xlsx
```

Next define the tabs, classes, and properties that you are interested in documenting.

e.g Tenants and Applications

```
tabs:
  # sheet name
  Tenants:
    # class query
    class: fvTenant
    # create a column for each of these properties
    properties: [name, dn]
```

Would produce the following sheet named `Tenants`

|name|dn|
|---|---|
|tenant1|uni/tn-tenant1|



You can optionally customize the header names for the sheet.  If not specified the column
name will be the property name.

```
  Applications:
    class: fvAp
    # optionally you may customize the header name
    headers: [Name, Distinguished Name]
    properties: [name, dn]  
```

Would produce the following sheet named `Applications`

|Name|Distinguished Name|
|---|---|
|app1|uni/tn-tenant1/ap-app1|

The [config.yml](./config.yml) also contains several other common classes as a starting point.  

## Running

Generating your spreadsheet is now a matter of running

    my-macbook-prompt ✗ python aci-doc.py
    Connected to ACI
    depending on your configuration, this could take a little while...
    Collecting information for Switches
    Collecting information for Subnets
    Collecting information for Applications
    Collecting information for Tenants
    Collecting information for Interface Errors
    Collecting information for Topology
