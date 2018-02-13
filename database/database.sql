

create table data_set (
    data_set_id        integer,
    data_set_name      varchar(40),
    description        varchar(1000)
);


create table datum (
    data_set_id        integer,
    datum_id           integer,
    frequency          double precision,
    tx_pol             char,
    rx_pol             char,
    tx_height          double precision,
    rx_height          double precision,
    tx_latitude        double precision,
    tx_longitude       double precision,
    rx_latitude        double precision,
    rx_longitude       double precision,
    distance           double precision,
    dbloss             double precision
);


create table profile_input (
    profile_input_id   integer,
    spacing            double precision,
    permittivity       double precision,
    conductivity       double precision
);


create table profile (
    data_set_id        integer,
    datum_id           integer,
    profile_input_id   integer
);


create table profile_point (
    data_set_id        integer,
    datum_id           integer,
    profile_input_id   integer,
    distance           double precision,
    height             double precision
);

